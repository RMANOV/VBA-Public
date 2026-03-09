// ═══════════════════════════════════════════════════════════════
// PATTERN: Text Split, Combine & Extract
// CATEGORY: Data Transformation | DIFFICULTY: Beginner-Intermediate
// ═══════════════════════════════════════════════════════════════
// WHAT: Parse, split, recombine, and extract parts of text
//       using M's built-in string functions — no regex needed.
// WHY:  Raw data comes with "Last, First M." names, mixed
//       formats, and embedded values. Clean parsing is essential.
// KEY LESSON: M has no regex, but BeforeDelimiter / AfterDelimiter
//             / Split covers 90% of parsing needs.
// ═══════════════════════════════════════════════════════════════

let
    // ── Sample data ─────────────────────────────────────────────
    Source = Table.FromRecords({
        [FullName = "Smith, John A.",     Email = "john.smith@acme.com",   Phone = "(555) 123-4567"],
        [FullName = "Doe, Jane",          Email = "jane.doe@widgets.org",  Phone = "555.987.6543"],
        [FullName = "O'Brien, Mary K.",   Email = "mary.ob@tech-corp.io",  Phone = "1-555-222-3333"],
        [FullName = "Garcia, Carlos R.",  Email = "cgarcia@startup.co",    Phone = "+1 (555) 444-5555"],
        [FullName = "Kim, Sarah",         Email = "s.kim@example.com",     Phone = "5550001234"]
    }),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 1: Text.BeforeDelimiter / Text.AfterDelimiter
    // ═══════════════════════════════════════════════════════════
    // Extract parts relative to a known separator.
    // These are the workhorses for structured text.

    // "Smith, John A." → LastName = "Smith", FirstAndMiddle = "John A."
    WithLastName = Table.AddColumn(
        Source,
        "LastName",
        // Everything BEFORE the first comma
        each Text.BeforeDelimiter([FullName], ","),
        type text
    ),

    WithFirstMiddle = Table.AddColumn(
        WithLastName,
        "FirstMiddle",
        // Everything AFTER the first comma, trimmed
        each Text.Trim(Text.AfterDelimiter([FullName], ",")),
        type text
    ),

    // Extract just the first name (before the space, if middle exists)
    WithFirstName = Table.AddColumn(
        WithFirstMiddle,
        "FirstName",
        each
            let
                fm = [FirstMiddle]
            in
                if Text.Contains(fm, " ") then
                    Text.BeforeDelimiter(fm, " ")
                else
                    fm,  // No middle initial — the whole thing is first name
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 2: Text.Split — break into a list by delimiter
    // ═══════════════════════════════════════════════════════════
    // Returns a list of parts. Access by index {0}, {1}, etc.
    // Useful when you have variable-length segments.

    WithNameParts = Table.AddColumn(
        WithFirstName,
        "NameParts",
        // Split "Smith, John A." on ", " → {"Smith", "John A."}
        each Text.Split([FullName], ", "),
        type list
    ),

    // Access individual parts by index
    WithSplitDemo = Table.AddColumn(
        WithNameParts,
        "SplitLast",
        each [NameParts]{0},  // First element = last name
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 3: Text.Combine — build strings from parts
    // ═══════════════════════════════════════════════════════════
    // Joins a list of text values with a separator.
    // The inverse of Text.Split.

    WithDisplayName = Table.AddColumn(
        WithSplitDemo,
        "DisplayName",
        // "John" + " " + "Smith" → "John Smith"
        each Text.Combine({[FirstName], [LastName]}, " "),
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 4: Table.SplitColumn — split a column into many
    // ═══════════════════════════════════════════════════════════
    // When you want to replace one column with multiple columns
    // in a single step. Each part becomes its own column.

    // Make a copy to demonstrate SplitColumn independently
    SplitDemo = Table.SplitColumn(
        Source,
        "FullName",
        Splitter.SplitTextByEachDelimiter({","}, QuoteStyle.None, false),
        {"Last", "FirstMiddle_Raw"}
    ),

    // Clean up whitespace on the split result
    SplitCleaned = Table.TransformColumns(
        SplitDemo,
        {{"FirstMiddle_Raw", Text.Trim, type text}}
    ),

    // ═══════════════════════════════════════════════════════════
    // PATTERN 5: Extract domain from email
    // ═══════════════════════════════════════════════════════════
    // Text.AfterDelimiter with "@" — simple and robust.

    WithDomain = Table.AddColumn(
        WithDisplayName,
        "EmailDomain",
        each Text.AfterDelimiter([Email], "@"),
        type text
    ),

    // Extract username too
    WithUsername = Table.AddColumn(
        WithDomain,
        "EmailUser",
        each Text.BeforeDelimiter([Email], "@"),
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // BONUS: Text.Select — keep only specific characters
    // ═══════════════════════════════════════════════════════════
    // Extracts ONLY characters that appear in the allowed list.
    // Perfect for cleaning phone numbers to digits only.

    WithCleanPhone = Table.AddColumn(
        WithUsername,
        "PhoneDigits",
        // Keep only digits 0-9, strip everything else
        each Text.Select([Phone], {"0".."9"}),
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // BONUS: Text.Remove — strip unwanted characters
    // ═══════════════════════════════════════════════════════════
    // The inverse of Text.Select — removes listed characters.

    WithStrippedPhone = Table.AddColumn(
        WithCleanPhone,
        "PhoneStripped",
        // Remove parentheses, dashes, spaces, dots, plus sign
        each Text.Remove([Phone], {"(", ")", "-", " ", ".", "+"}),
        type text
    ),

    // ═══════════════════════════════════════════════════════════
    // QUICK REFERENCE: Other useful text functions
    // ═══════════════════════════════════════════════════════════
    //
    // Text.Start(text, n)          → first n characters
    // Text.End(text, n)            → last n characters
    // Text.Middle(text, start, n)  → substring from position
    // Text.Range(text, start, n)   → same as Text.Middle
    // Text.Length(text)             → character count
    // Text.Contains(text, sub)     → true/false
    // Text.StartsWith(text, pre)   → true/false
    // Text.EndsWith(text, suf)     → true/false
    // Text.Replace(text, old, new) → replace substring
    // Text.Repeat(text, n)         → repeat n times
    // Text.Reverse(text)           → reverse characters
    // Text.Upper / Text.Lower      → case conversion
    // Text.Proper(text)            → Title Case
    // Text.Trim(text)              → strip leading/trailing whitespace
    // Text.Clean(text)             → remove non-printable characters
    // Text.PadStart(text, n, ch)   → left-pad to length n
    // Text.PadEnd(text, n, ch)     → right-pad to length n
    //
    // ═══════════════════════════════════════════════════════════

    // ── Final output — select the most useful columns ───────────
    Result = Table.SelectColumns(WithStrippedPhone, {
        "FullName", "FirstName", "LastName", "DisplayName",
        "Email", "EmailDomain", "EmailUser",
        "Phone", "PhoneDigits"
    })

in
    Result