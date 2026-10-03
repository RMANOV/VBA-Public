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
    //   condition — emit the current nonempty page, even if it is short
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
        // The condition tests the CURRENT candidate before the selector.
        // Emit its data even when it is the final short page.
        each [receivedCount] > 0,

        // ── NEXT STATE ──────────────────────────────────────────
        // A full page needs another request; a short page ends locally.
        each if [hasMore] then
            let
                nextOffset = [currentOffset] + PageSize,
                page       = FetchPage(nextOffset)
            in
                [
                    pageData      = page[data],
                    currentOffset = nextOffset,
                    receivedCount = List.Count(page[data]),
                    hasMore       = List.Count(page[data]) >= PageSize
                ]
        else
                [pageData = {}, currentOffset = [currentOffset],
                 receivedCount = 0, hasMore = false],

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
    // Empty input still needs the columns used by the typing step.
    // Keep ordinary record inference for nonempty data, including extra fields.
    Result = if List.IsEmpty(CombinedRecords) then
        #table(type table [id = Int64.Type, name = text, created_at = datetime], {})
    else
        Table.FromRecords(CombinedRecords),

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
    // the next request. Emit the current page before stopping on a null token.
    //
    // CursorPages = List.Generate(
    //     () =>
    //         let page = FetchWithCursor(null) in
    //         [
    //             pageData   = page[data],
    //             nextCursor = page[nextCursor],
    //             isTerminal = false
    //         ],
    //     each not [isTerminal],
    //     each if [nextCursor] <> null then
    //         let page = FetchWithCursor([nextCursor]) in
    //         [
    //             pageData   = page[data],
    //             nextCursor = page[nextCursor],
    //             isTerminal = false
    //         ]
    //     else [pageData = {}, nextCursor = null, isTerminal = true],
    //     each [pageData]
    // )
    //
    // ═══════════════════════════════════════════════════════════
    // VARIANT: LINK-BASED PAGINATION (RFC 5988)
    // ═══════════════════════════════════════════════════════════
    //
    // GitHub-style APIs return a Link header with rel="next".
    // Parse it to get the next URL; emit the current page before stopping.
    //
    // LinkPages = List.Generate(
    //     () =>
    //         let resp = Web.Contents(BaseUrl) in
    //         [
    //             pageData = Json.Document(resp),
    //             nextUrl  = GetNextLink(Value.Metadata(resp)),
    //             isTerminal = false
    //         ],
    //     each not [isTerminal],
    //     each if [nextUrl] <> null then
    //         let resp = Web.Contents([nextUrl]) in
    //         [
    //             pageData = Json.Document(resp),
    //             nextUrl  = GetNextLink(Value.Metadata(resp)),
    //             isTerminal = false
    //         ]
    //     else [pageData = {}, nextUrl = null, isTerminal = true],
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