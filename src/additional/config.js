/**
 * ===================================================================
 * CONFIGURATION
 * -------------------------------------------------------------------
 * All sheet names, ranges, and row/column counts are stored here.
 * ===================================================================
 */
const CONFIG = {
  // --- Environment ---
  // See src/additional/env.js. Any spreadsheet other than the one below is
  // treated as a DEV copy.
  ENV: {
    PROD_SPREADSHEET_ID: '1ANVfUER8MInJ_y-oRp7Nguwk0yDI9LtBt36PWBMmZbU',
    ALLOW_APIFY_IN_DEV: false // Flip to true only to deliberately spend credits from DEV
  },

  // --- Sheet Names ---
  SHEETS: {
    ARCHIVE_YEARS: [
      'Daily Archive 2026', 
      'Daily Archive 2025', 
      'Daily Archive 2024',
      'Daily Archive 2023'
    ],
    
    // The current active sheet for writing daily stats
    ARCHIVE: 'Daily Archive 2026', 
    
    TOOLS: 'Tools',
    LATEST: 'Latest',
    TRACKS: 'Tracks',
    MILESTONE_LOG: 'Milestone Log',
    ALBUMS: 'Albums',
    SONGS: 'Tracklist',
    COVERS: 'Covers',               // Cover key in column A
    CATEGORIES: 'Categories',       // One row per category (src/additional/categories.js)
    PENDING: 'Pending',             // Songs waiting to be added (src/songs/add_songs.js)
    TOTAL_ARCHIVE: 'Total Archive'  // Hand-kept cumulative backup; the script only reads it (checks)
  },

  // --- Health checks (see src/checks/checks.js) ---
  CHECKS: {
    MAX_DETAILS: 8,       // Problem lines shown per check in the report dialog
    SPARE_ROWS_WARN: 5,   // Warn when fewer spare category rows than this are left
    RAW_ROWS_WARN: 950    // Warn when the raw import gets this close to its 999-row range
  },

  // --- Songs registry (see src/additional/songs.js) ---
  // Header text of each column the code reads. Matching ignores case and spaces, so
  // "Track ID" and "trackId" are the same column; the column order does not matter.
  SONGS_SHEET: {
    HEADERS: {
      row: 'row',
      status: 'status',
      category: 'category',
      coverKey: 'coverKey',
      title: 'title',
      trackId: 'trackId'
    }
  },

  // --- Categories sheet (see src/additional/categories.js) ---
  // Found by header, ignoring case and spaces, like the Tracklist.
  CATEGORIES_SHEET: {
    HEADERS: {
      name: 'name',
      row: 'row',
      type: 'type',
      summaryCell: 'summaryCell',
      summaryLimit: 'summaryLimit'
    }
  },

  // --- Pending sheet (see src/songs/add_songs.js) ---
  PENDING_SHEET: {
    HEADERS: {
      title: 'title',
      category: 'category',
      coverKey: 'coverKey',
      trackId: 'trackId',
      status: 'status',
      spotifyTitle: 'Spotify Title',
      album: 'album',
      result: 'result'
    },
    DONE_PREFIX: '✅'   // A result starting with this marks the row as already added
  },

  // --- Row layout ---
  // A row means the same thing in Latest, Tools (J:M), every Daily Archive and the Total Archive:
  //   1       headers / dates
  //   2       Total Artist Streams
  //   3       Total Artist Solo Streams (the total minus SOLO_EXCLUDES)
  //   4-27    the categories (Categories sheet); up to LAST_AGGREGATE_ROW is spare for new ones
  //   50 ->   songs, open-ended. The last song row is the highest `row` in the Tracklist.
  LAYOUT: {
    TOTAL_ROW: 2,
    SOLO_ROW: 3,
    SOLO_EXCLUDES: 'Features',
    LAST_AGGREGATE_ROW: 49,
    FIRST_SONG_ROW: 50
  },

  // --- Range Definitions ---
  TOOLS: {
    RAW_DATA: 'E2:H1000',       // Raw import: E album, F Spotify name, G stream count, H track ID
    TOTALS_COLUMN: 12,          // Col L - each song's total, written by matchTotalsById()
    DAILY_COLUMN: 13,           // Col M - each song's daily (sheet formula)
    // Written into Col L for a song whose track ID is missing from the import. It breaks the
    // daily formula in Col M, so C1 becomes an error and main() refuses to run.
    MISSING_MARKER: '#MISSING',
    SUM_OF_DAILYS: 'C1'
  },
  LATEST: {
    DATE_CELL: 'Q1',            // Cell to increment date
    // Songs and albums share these columns.
    COLS: {
      TITLE: 5,                 // E
      TOTAL: 6,                 // F
      DAILY: 7,                 // G
      DAILY_PERCENT: 8,         // H
      DAILY_CHANGE: 10,         // J
      WEEKLY_CHANGE: 11,        // K
      BEST_SINCE: 12,           // L - written by findBestSince()
      DAY_AGO: 13,              // M - written by updateStats()
      WEEK_AGO: 14              // N - written by updateStats()
    }
  },
  ARCHIVE: {
    YESTERDAY_COLUMN: 3,        // Col C - newest first, so C is yesterday
    WEEK_AGO_COLUMN: 9          // Col I - a week ago
  },
  ALBUMS: {
    TOTAL_SUMMARY: 'G5'         // Albums sheet cell for the whole-discography summary
  },
  TRACKS: {
    UPCOMING_MILESTONES_DEST: 'N45',       // Start cell for header
    UPCOMING_MILESTONES_CLEAR: 'N45:Q549'  // Range to clear
  },

  // --- Apify Definitions ---
  APIFY: {
    ACTOR_ID: 'YZhD6hYc8daYSWXKs', 
    ALBUMS: {
      "Taylor Swift (Deluxe Edition)": "https://open.spotify.com/album/5eyZZoQEFQWRHkV2xgAeBw",
      "Fearless (Platinum Edition)": "https://open.spotify.com/album/2gP2LMVcIFgVczSJqn340t",
      "Speak Now (Deluxe Package)": "https://open.spotify.com/album/6S6JQWzUrJVcJLK4fi74Fw",
      "Red (Deluxe Edition)": "https://open.spotify.com/album/1KVKqWeRuXsJDLTW0VuD29",
      "1989 (Deluxe)": "https://open.spotify.com/album/1yGbNOtRIgdIiGHOEBaZWf",
      "reputation": "https://open.spotify.com/album/6DEjYFkNZh67HP7R9PSZvv",
      "Lover": "https://open.spotify.com/album/1NAmidJlEaVgA3MpcPFYGq",
      "folklore: the long pond studio sessions (from the Disney+ special) [deluxe edition]": "https://open.spotify.com/album/0PZ7lAru5FDFHuirTkWe9Z",
      "evermore (deluxe version)": "https://open.spotify.com/album/6AORtDjduMM3bupSWzbTSG",
      "Fearless (Taylor's Version)": "https://open.spotify.com/album/4hDok0OAJd57SGIT8xuWJH",
      "Red (Taylor's Version)": "https://open.spotify.com/album/6kZ42qRrzov54LcAk4onW9",
      "Midnights (The Til Dawn Edition)": "https://open.spotify.com/album/1fnJ7k0bllNfL1kVdNVW1A",
      "Speak Now (Taylor's Version)": "https://open.spotify.com/album/5AEDGbliTTfjOB8TSm1sxt",
      "1989 (Taylor's Version) [Deluxe]": "https://open.spotify.com/album/1o59UpKw81iHR0HPiSkJR0",
      "THE TORTURED POETS DEPARTMENT: THE ANTHOLOGY": "https://open.spotify.com/album/5H7ixXZfsNMGbIE5OBSpcb",
      "The Life of a Showgirl + Acoustic Collection": "https://open.spotify.com/album/6QNMhoV8V0u7cFuhhUBOn7",
      "Live From Clear Channel Stripped 2008": "https://open.spotify.com/album/1ycoesYxIFymXWebfmz828",
      "The Taylor Swift Holiday Collection": "https://open.spotify.com/album/7vzYp7FrKnTRoktBYsx9SF",
      "Speak Now World Tour Live": "https://open.spotify.com/album/6fyR4wBPwLHKcRtxgd4sGh",
      "Fearless (International Version)": "https://open.spotify.com/album/08CWGiv27MVQhYpuTtvx83",
      "The Man (Live From Paris)": "https://open.spotify.com/album/6l7iXnb2Y4yDR9zag3kckA",
      "Daylight (Live From Paris)": "https://open.spotify.com/album/2apcAEM6coXOMnHitrpRDk",
      "Death By A Thousand Cuts (Live From Paris)": "https://open.spotify.com/album/5nDpkszadFMGW6ZSYM9Q1V",
      "Lover (Live From Paris)": "https://open.spotify.com/album/7hvsmGyWH2kJS5X4E4t039",
      "ME! (Live From Paris)": "https://open.spotify.com/album/1jIp7CChnwdj9zUCvPxzQ7",
      "The Archer (Live From Paris)": "https://open.spotify.com/album/2A1msASmUbUdaZyeOxpbAD",
      "You Need To Calm Down (Live From Paris)": "https://open.spotify.com/album/1w1zVWd1JmsqAgfCw117Ra",
      "Cornelia Street (Live From Paris)": "https://open.spotify.com/album/4CF0YV0iNyfKoDt9jHbGj7",
      "Today Was A Fairytale": "https://open.spotify.com/album/7cfhaRiLqBzuAzc6Q24nyW",
      "The Hunger Games: Songs From District 12 And Beyond": "https://open.spotify.com/album/45nqVXRAW0xv0wpU9JljPN",
      "Hannah Montana The Movie": "https://open.spotify.com/album/1fc8tPf36cZhNYpNFrWh7o",
      "Sweeter Than Fiction": "https://open.spotify.com/album/11e4xCXllbvk8pWc1cCas1",
      "Fifty Shades Darker (Original Motion Picture Soundtrack)": "https://open.spotify.com/album/5VML6S956h4YfoYPooqLEi",
      "Cats: Highlights From The Motion Picture Soundtrack": "https://open.spotify.com/album/3OjtXmDVcfHAjJOS3xBoYU",
      "Only The Young (Featured in Miss Americana)": "https://open.spotify.com/album/5LGsh3kexUfi3qkIIxb8vK",
      "Carolina (From The Motion Picture “Where The Crawdads Sing”)": "https://open.spotify.com/album/5Bwg2XxrjTlrNy6BC7KQZf",
      "Love Drunk": "https://open.spotify.com/album/7DphDayDRJ1NtRmveflFWD",
      "Strange Clouds": "https://open.spotify.com/album/7qqCw47pAWFzhwTpVRd0zE",
      "Two Lanes Of Freedom": "https://open.spotify.com/album/08eQjFT0T2EGVFG5WeCe2V",
      "Bigger": "https://open.spotify.com/album/56XXDc04Gugu3CknMcsWLY",
      "How Long Do You Think It's Gonna Last?": "https://open.spotify.com/album/3YbMxdapL6mvSQjosFkc0T",
      "Renegade (Pop Version)": "https://open.spotify.com/album/49yiYOnz2UyaNl72xvDERt",
      "Women In Music Pt III (Expanded Edition)": "https://open.spotify.com/album/79thwyFL6Uo6rgTp3YWEAf",
      "The Joker And The Queen (feat. Taylor Swift)": "https://open.spotify.com/album/0vkAczpFKCazPKaoLtnBr0",
      "Live in No Shoes Nation": "https://open.spotify.com/album/2njb3cHa1yhUMdu8PT2VhY",
      "First Two Pages of Frankenstein": "https://open.spotify.com/album/5Mc6uebYtKnRc5I7bjlNB6",
      "The Secret of Us": "https://open.spotify.com/album/56bdWeO40o3WfAD2Lja4dl",
      "Ronan": "https://open.spotify.com/album/4T5606j6qpkQrWlwbKPLOp",
      "Christmas Tree Farm": "https://open.spotify.com/album/2jEFKhESvCLGpFP8KwpV2T",
      "All Of The Girls You Loved Before": "https://open.spotify.com/album/1Uauz6ql2dIPvIOH4JiuhD",
      "If This Was A Movie (Taylor’s Version)": "https://open.spotify.com/album/6IZm7NfvWyXp952VF36Z5F",
      "Eyes Open (Taylor's Version)": "https://open.spotify.com/album/6AeF9IkXFpHz5H3wUNX3L3",
      "Safe & Sound (Taylor's Version)": "https://open.spotify.com/album/4eZKjfeKPSwd6NYsmlKjuR",
      "You're Losing Me (From The Vault)": "https://open.spotify.com/album/5q3jthpn2h59P7pe2gmAl7",
      "Love Story (Digital Dog Remix)": "https://open.spotify.com/album/2Z2KdJE0nGGu0qdWA45mza",
      "Love Story (Pop Mix)": "https://open.spotify.com/album/1iab5rfjNpGhoPlFzPyp4k",
      "You're Not Sorry (CSI Remix)": "https://open.spotify.com/album/5sZIjREu3s225wwqJkgsYV",
      "Bad Blood (feat. Kendrick Lamar)": "https://open.spotify.com/album/1Tv3rrFNdXGtTeP1plX2xE",
      "Wildest Dreams (R3hab Remix)": "https://open.spotify.com/album/4vkDxA22dFzObjOY1nnXPk",
      "...Ready For It? (BloodPop® Remix)": "https://open.spotify.com/album/45fMDoh9dhhQicddIZzhKM",
      "Delicate (Sawyr And Ryan Tedder Mix)": "https://open.spotify.com/album/7HEXQkjBNiZilJcTWXwLOA",
      "Delicate (Seeb Remix)": "https://open.spotify.com/album/7gU675c7KZ54MzEcL1O3px",
      "Spotify Singles": "https://open.spotify.com/album/74utZeTCeaXy01BjOddyv8",
      "You Need To Calm Down (Clean Bandit Remix)": "https://open.spotify.com/album/531nfs5NPsmkRC1LJb1cdj",
      "Lover (First Dance Remix)": "https://open.spotify.com/album/6Ou4LWiU2Vu2V7KdHzw9At",
      "Lover (Remix) [feat. Shawn Mendes]": "https://open.spotify.com/album/2UfvnX1YYeC2cExMQTMbXC",
      "cardigan (cabin in candlelight version)": "https://open.spotify.com/album/3kqqalY92DENp7FiztDOjH",
      "betty (Live from the 2020 Academy of Country Music Awards)": "https://open.spotify.com/album/5smqkYeHq9jKRbaXpy3TpL",
      "the lakes (original version)": "https://open.spotify.com/album/40cMfQDrBCDmOaWZuNEmKq",
      "willow (90's trend remix)": "https://open.spotify.com/album/5C41iVpK8HXCe3qlLL3I38",
      "willow (the witch collection)": "https://open.spotify.com/album/6WzAiEDGTU7KmEyGwLpBXB",
      "Love Story (Taylor's Version) [Elvira Remix]": "https://open.spotify.com/album/3x4gaf5IPyFQNrxZY07CXA",
      "All Too Well (10 Minute Version) (The Short Film)": "https://open.spotify.com/album/2O1NYIBQCUobrL97A2Unk8",
      "All Too Well (Sad Girl Autumn Version) - Recorded at Long Pond Studios": "https://open.spotify.com/album/4qgs0gHJBgycj5SKqafFOB",
      "Message In A Bottle (Fat Max G Remix) (Taylor’s Version)": "https://open.spotify.com/album/6d8IUfMwq7HGCnR2efXjdm",
      "Anti-Hero (Acoustic Version)": "https://open.spotify.com/album/5vgWXUueeEU2labRz6TlNv",
      "Anti-Hero (ILLENIUM Remix)": "https://open.spotify.com/album/20wq0dFrgEhhulGXqnb4A6",
      "Anti-Hero (Remixes)": "https://open.spotify.com/album/7irmI5g3OLC1gUXlxysOWt",
      "Lavender Haze (Acoustic Version)": "https://open.spotify.com/album/6eKdbTio5viiwJ5FE5J8wU",
      "Lavender Haze (Remixes)": "https://open.spotify.com/album/5LyzI39gkePgpHz38bEQIr",
      "The Cruelest Summer": "https://open.spotify.com/album/12A1Byk8EpqzaHSw12nKyW",
      "Fortnight (feat. Post Malone) [BLOND:ISH Remix]": "https://open.spotify.com/album/1agEHpWnELaZHWzcbGDCpu",
      "Fortnight (Acoustic Version)": "https://open.spotify.com/album/5IYMYmX28qpI6OEnQhdazX",
      "I Can Do It With a Broken Heart (Dombresky Remix)": "https://open.spotify.com/album/3WNGsnTetig4bJbw2BMbK7",
      "The Fate of Ophelia (Loud Luxury Remix)": "https://open.spotify.com/album/1Ed98OfVXz3CiJcupzuLTo",
      "The Life of a Showgirl (Track by Track Version)": "https://open.spotify.com/album/4tFsJC9jQ0Qjj7HSExhAdm",
      "The Fate of Ophelia (The Chainsmokers Remix)": "https://open.spotify.com/album/264e9sevSUiitcQeM1y1C5",
      "Opalite [Official Music Video (Extended Versions)]":"https://open.spotify.com/album/0VBjFJFcHhu7QDocog6D3s?si=0qgg6FUjTwyQ5zoLxEq_UQ",
      "Opalite - BUNT. Remix": "https://open.spotify.com/album/5oO9uwBHZAawqYnmPpVUUi",
      "Opalite - Ely Oaks Remix": "https://open.spotify.com/album/72hPyRss6soJeGQLEek1qr",
      "Opalite - Skream Remix": "https://open.spotify.com/album/6WNAGQ9MfrSyOFbMc66pCz?si=0DYd2aC9Qeu_dSS5knZKDg",
      "Opalite - Chris Lake Remix": "https://open.spotify.com/album/1OioQUIAwZ6PE2V7UJUHX2",
      "I Knew It,I Knew You": "https://open.spotify.com/album/4Ii9whWXI1O1H01ziECRaG"
      //"Red (Big Machine Radio Release Special)": "https://open.spotify.com/album/4jTYApZPMapg56gRycOn0D?si=cq650LrrQS64DnNoFnSZFA"
    }
  }
};