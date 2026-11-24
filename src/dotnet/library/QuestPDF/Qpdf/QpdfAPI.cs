using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using QuestPDF.Fluent;
using QuestPDF.Skia;

namespace QuestPDF.Qpdf;

static class QpdfAPI
{
    public static int GetCompatibilityVersion()
    {
        return API.questpdf_get_compatibility_version();
    }
    
    public static void ExecuteJob(string jobJson, Func<byte[], byte[]>? transformMetadata = null,
        IReadOnlyList<DocumentOperation.Bookmark>? bookmarks = null, Action? beforeWrite = null, bool replaceBookmarks = false)
    {
        WithDocument(jobJson, (job, document, checkResult) =>
        {
            if (transformMetadata != null)
                TransformMetadata(job, document, transformMetadata, checkResult);

            if (replaceBookmarks || bookmarks is { Count: > 0 })
                WriteBookmarks(document, bookmarks ?? Array.Empty<DocumentOperation.Bookmark>(), replaceBookmarks);

            beforeWrite?.Invoke();
            checkResult(API.qpdfjob_write_qpdf(job, document));
        });
    }

    public static DocumentOperation.DocumentInfo InspectFile(string filePath, string? password,
        string? pageSelector = null, bool readBookmarks = true)
    {
        var configuration = new JobConfiguration
        {
            InputFile = filePath,
            Password = password,
            Pages = pageSelector == null
                ? null
                : new[] { new JobConfiguration.PageConfiguration { File = ".", Range = pageSelector } },
            ShowNpages = string.Empty
        };

        var pageCount = 0;
        using var output = readBookmarks ? new MemoryStream() : null;
        WithDocument(QpdfJobSerializer.Serialize(configuration), (_, document, _) =>
        {
            pageCount = ReadPageCount(document);

            if (output != null)
            {
                // Keep the real input path separate from the JSON writer's stdout marker.
                const string outlineJob = """{"empty":"","outputFile":"-","json":"2","jsonKey":["outlines"]}""";
                WithJob(outlineJob, (writer, checkResult) => checkResult(API.qpdfjob_write_qpdf(writer, document)), output);
            }
        });

        if (output == null)
            return new DocumentOperation.DocumentInfo(pageCount, Array.Empty<DocumentOperation.Bookmark>());

        var buffer = output.GetBuffer();
        var length = (int)output.Length;
        var utf8 = new ReadOnlySpan<byte>(buffer, 0, length);
        if (utf8.IndexOf("\uFFFE"u8) >= 0 || utf8.IndexOf("\uFFFF"u8) >= 0)
        {
            // The framework JSON reader accepts these Unicode characters only in escaped form.
            var json = Encoding.UTF8.GetString(buffer, 0, length).Replace("\uFFFE", "\\uFFFE").Replace("\uFFFF", "\\uFFFF");
            buffer = Encoding.UTF8.GetBytes(json);
            length = buffer.Length;
        }

        using var reader = JsonReaderWriterFactory.CreateJsonReader(buffer, 0, length, XmlDictionaryReaderQuotas.Max);
        var outlines = XElement.Load(reader, LoadOptions.PreserveWhitespace).Element("outlines");
        return new DocumentOperation.DocumentInfo(pageCount, ReadBookmarks(outlines));
    }

    private static IReadOnlyList<DocumentOperation.Bookmark> ReadBookmarks(XElement? outlines)
    {
        if (outlines == null)
            return Array.Empty<DocumentOperation.Bookmark>();

        var bookmarks = new List<DocumentOperation.Bookmark>();
        foreach (var item in outlines.Elements("item"))
        {
            var title = item.Element("title")?.Value ?? string.Empty;
            var destination = item.Element("destpageposfrom1");
            var pageNumber = destination?.Attribute("type")?.Value == "number" ? (int?)destination : null;
            var children = ReadBookmarks(item.Element("kids"));
            bookmarks.Add(new DocumentOperation.Bookmark(title, pageNumber, children));
        }

        return bookmarks;
    }

    private static void WithDocument(string jobJson, Action<IntPtr, IntPtr, Action<int>> process)
    {
        WithJob(jobJson, (job, checkResult) =>
        {
            var document = API.qpdfjob_create_qpdf(job);
            if (document == IntPtr.Zero)
                checkResult(JobResultError);

            try
            {
                process(job, document, checkResult);
            }
            finally
            {
                API.qpdf_cleanup(ref document);
            }
        });
    }

    private static void WithJob(string jobJson, Action<IntPtr, Action<int>> process, MemoryStream? output = null)
    {
        QpdfNativeDependencyCompatibilityChecker.Test();

        using var error = new MemoryStream();
        var errorHandle = GCHandle.Alloc(error);
        var outputHandle = default(GCHandle);

        var logger = IntPtr.Zero;
        var jobHandle = IntPtr.Zero;

        try
        {
            // create logger
            logger = API.qpdflogger_create();
            API.qpdflogger_set_error(logger, CustomLogDestination, WriteCallbackPointer, GCHandle.ToIntPtr(errorHandle));

            if (output != null)
            {
                outputHandle = GCHandle.Alloc(output);
                API.qpdflogger_set_save(logger, CustomLogDestination, WriteCallbackPointer, GCHandle.ToIntPtr(outputHandle), 0);
            }

            // prepare the job
            jobHandle = API.qpdfjob_init();
            API.qpdfjob_set_logger(jobHandle, logger);
            ThrowOnJobError(API.qpdfjob_initialize_from_json(jobHandle, jobJson));

            process(jobHandle, ThrowOnJobError);
        }
        finally
        {
            if (jobHandle != IntPtr.Zero)
                API.qpdfjob_cleanup(ref jobHandle);

            if (logger != IntPtr.Zero)
                API.qpdflogger_cleanup(ref logger);

            if (outputHandle.IsAllocated)
                outputHandle.Free();

            errorHandle.Free();
        }

        void ThrowOnJobError(int jobResultId)
        {
            // 0 = success, 1 = undefined, 2 = error, 3 = warning
            if (jobResultId != JobResultError)
                return;

            var message = Encoding.UTF8.GetString(error.GetBuffer(), 0, (int)error.Length);
            throw new Exception($"QuestPDF could not perform document operation:\n\n{message}");
        }
    }

    private static void TransformMetadata(IntPtr job, IntPtr document, Func<byte[], byte[]> transform, Action<int> checkResult)
    {
        checkResult(API.questpdf_job_get_xmp_metadata(job, document, out var buffer, out var length));

        if (buffer == IntPtr.Zero)
        {
            throw new Exception(
                "QuestPDF could not extend the document metadata because the document does not contain any XMP metadata. " +
                "If the document is generated with QuestPDF, please enable the PDF/A or PDF/UA conformance in the DocumentSettings.");
        }

        byte[] content;

        try
        {
            content = new byte[(int)length];
            Marshal.Copy(buffer, content, 0, content.Length);
        }
        finally
        {
            API.qpdf_oh_free_buffer(ref buffer);
        }

        content = transform(content);

        // qpdf copies the content, so the managed array is needed only for the duration of the call
        checkResult(API.questpdf_job_set_xmp_metadata(job, document, content, (UIntPtr)content.Length));
    }

    private static int ReadPageCount(IntPtr document)
    {
        API.qpdf_silence_errors(document);
        var pageCount = API.qpdf_get_num_pages(document);
        ThrowOnDocumentError(document);
        return pageCount;
    }

    private static void WriteBookmarks(IntPtr document, IReadOnlyList<DocumentOperation.Bookmark> bookmarks, bool replace)
    {
        var pageCount = ReadPageCount(document);
        ValidatePageNumbers(bookmarks);

        var catalog = API.qpdf_get_root(document);
        if (replace)
        {
            API.qpdf_oh_remove_key(document, catalog, "/Outlines");
            ThrowOnDocumentError(document);

            if (bookmarks.Count == 0)
                return;
        }

        var outline = GetKey(catalog, "/Outlines");

        if (IsNull(outline))
        {
            outline = NewIndirectObject("<< /Type /Outlines >>");
            SetKey(catalog, "/Outlines", outline);
        }

        ThrowOnDocumentError(document);

        if (!IsIndirectDictionary(outline))
            throw new InvalidOperationException("The PDF contains an invalid outline dictionary.");

        var first = GetKey(outline, "/First");
        var last = GetKey(outline, "/Last");
        if (IsNull(first) != IsNull(last) ||
            (!IsNull(last) && (!IsIndirectDictionary(first) || !IsIndirectDictionary(last) || !IsNull(GetKey(last, "/Next")))))
            throw new InvalidOperationException("The PDF contains an invalid outline.");

        var visibleCount = GetVisibleCount() + AppendBookmarks(outline, last, bookmarks);
        SetKey(outline, "/Count", API.qpdf_oh_new_integer(document, visibleCount));
        ThrowOnDocumentError(document);

        void ValidatePageNumbers(IReadOnlyList<DocumentOperation.Bookmark> entries)
        {
            foreach (var entry in entries)
            {
                if (entry.PageNumber is int pageNumber && (pageNumber < 1 || pageNumber > pageCount))
                    throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, $"The output document has {pageCount} pages.");

                ValidatePageNumbers(entry.Children);
            }
        }

        long AppendBookmarks(uint parent, uint previous, IReadOnlyList<DocumentOperation.Bookmark> entries)
        {
            var added = 0L;

            foreach (var entry in entries)
            {
                var item = NewIndirectObject(entry.PageNumber.HasValue ? "<< /Dest [null /Fit] >>" : "<< >>");
                var title = Encoding.UTF8.GetBytes(entry.Title);
                SetKey(item, "/Title", API.qpdf_oh_new_binary_unicode_string(document, title, (UIntPtr)title.Length));
                SetKey(item, "/Parent", parent);

                if (entry.PageNumber is int pageNumber)
                {
                    var page = API.qpdf_get_page_n(document, (UIntPtr)(pageNumber - 1));
                    API.qpdf_oh_set_array_item(document, GetKey(item, "/Dest"), 0, page);
                }

                if (IsNull(previous))
                    SetKey(parent, "/First", item);
                else
                {
                    SetKey(previous, "/Next", item);
                    SetKey(item, "/Prev", previous);
                }

                if (entry.Children.Count > 0)
                {
                    var descendants = AppendBookmarks(item, GetKey(item, "/Last"), entry.Children);
                    SetKey(item, "/Count", API.qpdf_oh_new_integer(document, descendants));
                    added += descendants;
                }

                added++;
                previous = item;
            }

            SetKey(parent, "/Last", previous);
            return added;
        }

        uint NewIndirectObject(string value) => API.qpdf_make_indirect_object(document, API.qpdf_oh_parse(document, value));
        uint GetKey(uint item, string key) => API.qpdf_oh_get_key(document, item, key);
        bool IsNull(uint item) => API.qpdf_oh_is_null(document, item) != 0;
        bool IsIndirectDictionary(uint item) => API.qpdf_oh_get_object_id(document, item) > 0 && API.qpdf_oh_is_dictionary(document, item) != 0;
        void SetKey(uint item, string key, uint value) => API.qpdf_oh_replace_key(document, item, key, value);

        long GetVisibleCount()
        {
            if (IsNull(first))
                return 0;

            if (TryGetCount(outline, out var count) && count > 0)
                return count;

            // A missing root count can be reconstructed from the expanded top-level items.
            var visible = 0L;
            var visited = new HashSet<int>();
            for (var item = first; !IsNull(item); item = GetKey(item, "/Next"))
            {
                var id = API.qpdf_oh_get_object_id(document, item);
                if (id == 0 || !visited.Add(id))
                    throw new InvalidOperationException("The PDF contains an invalid outline.");

                visible++;
                if (TryGetCount(item, out var childCount) && childCount > 0)
                    visible += childCount;
            }

            return visible;
        }

        bool TryGetCount(uint item, out long count) =>
            API.qpdf_oh_get_value_as_longlong(document, GetKey(item, "/Count"), out count) != 0;
    }

    private static unsafe void ThrowOnDocumentError(IntPtr document)
    {
        var error = API.qpdf_get_error(document);
        if (error == IntPtr.Zero)
            return;

        var message = (byte*)API.qpdf_get_error_full_text(document, error);
        var length = 0;
        while (message[length] != 0)
            length++;

        throw new InvalidOperationException(Encoding.UTF8.GetString(message, length));
    }
    
    private const int JobResultError = 2;
    private const int CustomLogDestination = 4;

    #region Logging
    
    private static int WriteCallback(IntPtr data, UIntPtr length, IntPtr context)
    {
        try
        {
            var stream = (MemoryStream)GCHandle.FromIntPtr(context).Target!;
            var count = checked((int)length.ToUInt64());
            var bytes = new byte[count];
            Marshal.Copy(data, bytes, 0, count);
            stream.Write(bytes, 0, count);
            return 0;
        }
        catch
        {
            // Report callback failures to QPDF without letting managed exceptions cross into native code.
            return 1;
        }
    }
    
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    private delegate int CallbackDelegate(IntPtr data, UIntPtr length, IntPtr context);
    
    private static readonly CallbackDelegate WriteCallbackDelegate = WriteCallback;
    
    private static readonly IntPtr WriteCallbackPointer = Marshal.GetFunctionPointerForDelegate(WriteCallbackDelegate);
    
    #endregion
    
    private static class API
    {
        const string LibraryName = "questpdf_qpdf";
        
        /* GENERAL */
        
        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int questpdf_get_compatibility_version();
    
        /* JOBS */
        
        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr qpdfjob_init();
    
        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void qpdfjob_cleanup(ref IntPtr jobHandle);
    
        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int qpdfjob_initialize_from_json(IntPtr jobHandle, [MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(Utf8StringMarshaller))] string json);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr qpdfjob_create_qpdf(IntPtr jobHandle);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int qpdfjob_write_qpdf(IntPtr jobHandle, IntPtr documentHandle);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void qpdf_cleanup(ref IntPtr documentHandle);

        /* METADATA */

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int questpdf_job_get_xmp_metadata(IntPtr jobHandle, IntPtr documentHandle, out IntPtr buffer, out UIntPtr length);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int questpdf_job_set_xmp_metadata(IntPtr jobHandle, IntPtr documentHandle, byte[] content, UIntPtr length);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void qpdf_oh_free_buffer(ref IntPtr buffer);

        /* BOOKMARKS */

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void qpdf_silence_errors(IntPtr document);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr qpdf_get_error(IntPtr document);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr qpdf_get_error_full_text(IntPtr document, IntPtr error);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int qpdf_get_num_pages(IntPtr document);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint qpdf_get_page_n(IntPtr document, UIntPtr pageIndex);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint qpdf_get_root(IntPtr document);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint qpdf_make_indirect_object(IntPtr document, uint item);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint qpdf_oh_parse(IntPtr document, string value);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int qpdf_oh_is_null(IntPtr document, uint item);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int qpdf_oh_is_dictionary(IntPtr document, uint item);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int qpdf_oh_get_object_id(IntPtr document, uint item);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint qpdf_oh_get_key(IntPtr document, uint item, string key);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void qpdf_oh_replace_key(IntPtr document, uint item, string key, uint value);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void qpdf_oh_remove_key(IntPtr document, uint item, string key);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int qpdf_oh_get_value_as_longlong(IntPtr document, uint item, out long value);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint qpdf_oh_new_integer(IntPtr document, long value);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint qpdf_oh_new_binary_unicode_string(IntPtr document, byte[] value, UIntPtr length);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void qpdf_oh_set_array_item(IntPtr document, uint array, int index, uint item);
        
        /* LOGGING */
        
        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = nameof(qpdflogger_create))]
        public static extern IntPtr qpdflogger_create();

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = nameof(qpdflogger_cleanup))]
        public static extern void qpdflogger_cleanup(ref IntPtr loggerHandle);
        
        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = nameof(qpdflogger_set_error))]
        public static extern void qpdflogger_set_error(IntPtr loggerHandle, int destination, IntPtr callBackHandler, IntPtr udata);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void qpdflogger_set_save(IntPtr loggerHandle, int destination, IntPtr callback, IntPtr context, int onlyIfNotSet);
        
        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, EntryPoint = nameof(qpdfjob_set_logger))]
        public static extern void qpdfjob_set_logger(IntPtr jobHandle, IntPtr loggerHandle);
    }
}
