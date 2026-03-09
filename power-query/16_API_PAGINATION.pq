// ═══════════════════════════════════════════════════════════════
// PATTERN: API Pagination with List.Generate
// CATEGORY: Data Acquisition | DIFFICULTY: Advanced
// ═══════════════════════════════════════════════════════════════
// WHAT: Consume a paginated REST API by looping until all pages
//       are fetched, then combine into a single table.
// WHY:  Most APIs return data in pages (50-1000 records each).
//       You must loop through all pages to get the full dataset.
// KEY LESSON: API pagination is List.Generate's killer use case.
//             Works with offset-based, cursor-based, and
//             link-based pagination — just swap the "next" logic.
// ═══════════════════════════════════════════════════════════════

let
    // ── Configuration ───────────────────────────────────────────
    BaseUrl  = "https://api.example.com/v1/records",
    PageSize = 100,  // Records per page (match API's max)

    // ── Helper: Fetch a single page ─────────────────────────────
    // Encapsulates the Web.Contents call so each page request
    // is clean and consistent. Add headers (API keys, auth
    // tokens) here once — all pages inherit them.
    FetchPage = (offset as number) as record =>
        let
            response = Web.Contents(
                BaseUrl,
                [
                    // Query parameters appended to URL:
                    // ?offset=0&limit=100, ?offset=100&limit=100, etc.
                    Query = [
                        offset = Text.From(offset),
                        limit  = Text.From(PageSize)
                    ],
                    Headers = [
                        #"Content-Type" = "application/json",
                        #"Accept"       = "application/json"
                        // Add your API key here:
                        // , #"Authorization" = "Bearer " & ApiKey
                    ],
                    // Prevent Power Query from caching API responses
                    // (important for real-time data)
                    IsRetry = true
                ]
            ),
            parsed = Json.Document(response)
        in
            parsed,

    // ═══════════════════════════════════════════════════════════
    // CORE: List.Generate drives the pagination loop
    // ═══════════════════════════════════════════════════════════
    //
    // List.Generate takes 4 arguments:
    //   initial   — starting state (page 0)
    //   condition — keep going? (did we get a full page?)
    //   next      — advance state (increment offset)
    //   selector  — what to extract from each state (the data)
    //
    // It produces a list of results, one entry per page.

    AllPages = List.Generate(
        // ── INITIAL STATE ───────────────────────────────────────
        // Fetch page 0 and build the initial state record.
        () =>
            let
                page = FetchPage(0)
            in
                [
                    pageData      = page[data],          // The array of records
                    currentOffset = 0,                   // Track position
                    receivedCount = List.Count(page[data]),  // How many came back
                    hasMore       = List.Count(page[data]) >= PageSize  // Full page = more data
                ],

        // ── CONDITION ───────────────────────────────────────────
        // Continue as long as the previous page was full.
        // A page with fewer records than PageSize means we've
        // reached the last page.
        each [hasMore],

        // ── NEXT STATE ──────────────────────────────────────────
        // Increment offset by PageSize, fetch next page, update flags.
        each
            let
                nextOffset = [currentOffset] + PageSize,
                page       = FetchPage(nextOffset)
            in
                [
                    pageData      = page[data],
                    currentOffset = nextOffset,
                    receivedCount = List.Count(page[data]),
                    hasMore       = List.Count(page[data]) >= PageSize
                ],

        // ── SELECTOR ────────────────────────────────────────────
        // Extract just the data array from each state.
        // This is what ends up in the final list.
        each [pageData]
    ),

    // ═══════════════════════════════════════════════════════════
    // COMBINE: Flatten all pages into one list of records
    // ═══════════════════════════════════════════════════════════
    // AllPages is a list of lists: {{rec1, rec2, ...}, {rec101, rec102, ...}, ...}
    // List.Combine merges them into a single flat list.
    CombinedRecords = List.Combine(AllPages),

    // ── Convert to table ────────────────────────────────────────
    Result = Table.FromRecords(CombinedRecords),

    // ── Type the columns (adjust to match your actual API) ──────
    Typed = Table.TransformColumnTypes(Result, {
        {"id", Int64.Type},
        {"name", type text},
        {"created_at", type datetime}
    })

    // ═══════════════════════════════════════════════════════════
    // VARIANT: CURSOR-BASED PAGINATION
    // ═══════════════════════════════════════════════════════════
    //
    // Some APIs use a cursor token instead of offset/limit.
    // The response includes a "nextCursor" field; pass it in
    // the next request. Loop until nextCursor is null.
    //
    // CursorPages = List.Generate(
    //     () =>
    //         let page = FetchWithCursor(null) in
    //         [
    //             pageData   = page[data],
    //             nextCursor = page[nextCursor],
    //             hasMore    = page[nextCursor] <> null
    //         ],
    //     each [hasMore],
    //     each
    //         let page = FetchWithCursor([nextCursor]) in
    //         [
    //             pageData   = page[data],
    //             nextCursor = page[nextCursor],
    //             hasMore    = page[nextCursor] <> null
    //         ],
    //     each [pageData]
    // )
    //
    // ═══════════════════════════════════════════════════════════
    // VARIANT: LINK-BASED PAGINATION (RFC 5988)
    // ═══════════════════════════════════════════════════════════
    //
    // GitHub-style APIs return a Link header with rel="next".
    // Parse it to get the full URL for the next page.
    //
    // LinkPages = List.Generate(
    //     () =>
    //         let resp = Web.Contents(BaseUrl) in
    //         [
    //             pageData = Json.Document(resp),
    //             nextUrl  = GetNextLink(Value.Metadata(resp)),
    //             hasMore  = GetNextLink(Value.Metadata(resp)) <> null
    //         ],
    //     each [hasMore],
    //     each
    //         let resp = Web.Contents([nextUrl]) in
    //         [
    //             pageData = Json.Document(resp),
    //             nextUrl  = GetNextLink(Value.Metadata(resp)),
    //             hasMore  = GetNextLink(Value.Metadata(resp)) <> null
    //         ],
    //     each [pageData]
    // )
    //
    // ═══════════════════════════════════════════════════════════
    // NOTES ON PRODUCTION USE
    // ═══════════════════════════════════════════════════════════
    //
    // 1. RATE LIMITING: Add Function.InvokeAfter for delay:
    //    Function.InvokeAfter(() => FetchPage(n), #duration(0,0,0,1))
    //
    // 2. ERROR HANDLING: Wrap Web.Contents in try...otherwise
    //    to handle 429 (rate limit) or 500 (server error).
    //
    // 3. CREDENTIALS: Use Power Query's built-in credential
    //    manager (Edit Credentials in Data Source Settings)
    //    rather than hardcoding API keys.
    //
    // 4. TIMEOUT: Web.Contents supports [Timeout = #duration()]
    //    for slow APIs.
    //
    // ═══════════════════════════════════════════════════════════

in
    Typed