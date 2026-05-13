/**
 * ===================================================================
 * CONFIGURATION
 * -------------------------------------------------------------------
 * All sheet names, ranges, and row/column counts are stored here.
 * ===================================================================
 */
const CONFIG = {
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
    ALBUMS: 'Albums'
  },
  
  // --- Row & Count Definitions ---
  SONGS: {
    START_ROW: 2,
    COUNT: 548 // e.g., rows 2 to 549
  },
  ALBUMS: {
    START_ROW: 550, // e.g., rows 550 to 574
    COUNT: 25,
    TOTAL_SUMMARY: 'G5',
  },

  // --- Range Definitions (A1 Notation) ---
  TOOLS: {
    RAW_DATA: 'F2:G1000',       // Where the raw imported data goes into
    SONG_DATA: 'L2:M549',       // Data to copy to Latest (Totals + Daily)
    DAILY_STREAMS: 'M2:M549',   // Daily streams to copy to Archive
    SUM_OF_DAILYS: 'C1'
  },
  LATEST: {
    DATE_CELL: 'Q1',           // Cell to increment date
    SONG_DATA_DEST: 'F2:G549', // Destination for TOOLS_SONG_DATA
    YESTERDAY_STREAMS_DEST: 'M2:M549',
    WEEK_STREAMS_DEST: 'N2:N549',
    SONG_BEST_SINCE_DEST: 'L2', // Output for "Best Since" (Songs)
    ALBUM_BEST_SINCE_DEST: 'Y2', // Output for "Best Since" (Albums)
    SONG_NAMES: 'E2:E549',
    ALBUMS_START_ROW: 2, 
    OVERALL_ROW: 26
  },
  ARCHIVE: {
    DAILY_STREAMS_DEST: 'B2:B549',
    YESTERDAY_STREAMS: 'C2:C549',
    WEEK_DATA: 'I2:I549'
  },
  TRACKS: {
    UPCOMING_MILESTONES_DEST: 'N45',       // Start cell for header
    UPCOMING_MILESTONES_CLEAR: 'N45:Q549'  // Range to clear
  },
  STATS: {
    'Taylor Swift (Debut)':                 { songRow: 2,   count: 15, summaryRow: 2, destCell: 'F23' },
    'Fearless (2008)':                      { songRow: 17,  count: 19, summaryRow: 3, destCell: 'F79' },
    'Speak Now (2010)':                     { songRow: 36,  count: 17, summaryRow: 4, destCell: 'L93' },
    'Red (2012)':                           { songRow: 57,  count: 19, summaryRow: 5, destCell: 'L119'},
    '1989 (2014)':                          { songRow: 79,  count: 16, summaryRow: 6, destCell: 'R23' },
    'reputation':                           { songRow: 98,  count: 15, summaryRow: 7, destCell: 'R78' },
    'Lover':                                { songRow: 113, count: 18, summaryRow: 8, destCell: 'W23' },
    'folklore':                             { songRow: 131, count: 34, summaryRow: 9, destCell: 'W51' },
    'evermore':                             { songRow: 165, count: 17, summaryRow: 10, destCell: 'W78' },
    "Fearless (Taylor's Version)":          { songRow: 182, count: 26, summaryRow: 13, destCell: 'F45' },
    "Red (Taylor's Version)":               { songRow: 208, count: 30, summaryRow: 14, destCell: 'L54' },
    'Midnights':                            { songRow: 238, count: 23, summaryRow: 11, destCell: 'W102'},
    "Speak Now (Taylor's Version)":         { songRow: 261, count: 22, summaryRow: 15, destCell: 'L23' },
    "1989 (Taylor's Version)":              { songRow: 283, count: 22, summaryRow: 16, destCell: 'R49' },
    'The Tortured Poets Department ':       { songRow: 305, count: 31, summaryRow: 12, destCell: 'AB23' },
    'The Life of a Showgirl':               { songRow: 336, count: 12, summaryRow: 17, destCell: 'AB89'}
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
      "Opalite - Chris Lake Remix": "https://open.spotify.com/album/1OioQUIAwZ6PE2V7UJUHX2"
      //"Red (Big Machine Radio Release Special)": "https://open.spotify.com/album/4jTYApZPMapg56gRycOn0D?si=cq650LrrQS64DnNoFnSZFA"
    }
  }
};