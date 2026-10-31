const SONGBOOK_SHEET_ID = '1KGqgpmBau3bmoWjfA0Hwa16p6Ggo8x9Vu8djrsyqno4';
const SONG_SHEET_NAME = 'Library';
const REQUEST_SHEET_NAME = 'Submitted Requests';

// Cached data structures
let cachedArtists = null;
let cachedSongsByArtist = null;

/**
 * Clean and normalize text from the sheet.
 * Fixes hidden characters, unicode issues, weird whitespace, etc.
 */
function clean(str) {
  return String(str)
    .normalize("NFKC")              // normalize unicode
    .replace(/\s+/g, " ")           // collapse weird whitespace
    .replace(/[\u200B-\u200D]/g, "")// remove zero-width chars
    .replace(/\u00A0/g, " ")        // remove non-breaking spaces
    .trim();
}

/**
 * Load and index the sheet ONCE.
 * Builds:
 *  - cachedArtists: sorted list of unique artists
 *  - cachedSongsByArtist: { artist: [ {title}, ... ] }
 */
function loadIndex() {
  if (cachedArtists && cachedSongsByArtist) {
    return; // Already indexed
  }

  const ss = SpreadsheetApp.openById(SONGBOOK_SHEET_ID);
  const sheet = ss.getSheetByName(SONG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const songsByArtist = {};
  const artistsSet = new Set();

  for (let i = 1; i < data.length; i++) {
    const rawArtist = data[i][0];
    const rawTitle = data[i][1];

    const artist = clean(rawArtist);
    const title = clean(rawTitle);

    if (!artist || !title) continue;

    if (!songsByArtist[artist]) {
      songsByArtist[artist] = [];
    }

    songsByArtist[artist].push({ artist, title });
    artistsSet.add(artist);
  }

  cachedArtists = Array.from(artistsSet).sort((a, b) =>
  a.toLowerCase().localeCompare(b.toLowerCase())
);
  cachedSongsByArtist = songsByArtist;
}

/**
 * Return a page of artists (50 at a time)
 */
function getArtistsPage(page, pageSize) {
  loadIndex();
  const start = page * pageSize;
  const end = start + pageSize;
  return cachedArtists.slice(start, end);
}

/**
 * Return songs for a single artist instantly
 */
function getSongsForArtist(artist) {
  loadIndex();
  return cachedSongsByArtist[artist] || [];
}

function getSongsForArtists(artists) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SONG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const result = {};

  // Initialize empty arrays for each artist
  artists.forEach(a => result[a] = []);

  // Fill arrays
  data.forEach(row => {
    const title = row[0];
    const artist = row[1];
    if (result[artist]) {
      result[artist].push({ title, artist });
    }
  });

  return result;
}


/**
 * Return a page of songs 50 at a time
 */
function getSongsPage(pageIndex, pageSize) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SONG_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  // Skip header row
  const rows = data.slice(1);

  // Clean + map
  const songs = rows.map(r => ({
    artist: clean(r[0]),
    title:  clean(r[1])
  }));

  // ⭐ FULL SORT — BEFORE paging
  songs.sort((a, b) =>
    a.title.toLowerCase().localeCompare(b.title.toLowerCase())
  );

  // Paging AFTER sorting
  const start = pageIndex * pageSize;
  const end   = start + pageSize;

  return songs.slice(start, end);
}



/** 
 * Add Starts with Letter option to Mode Search 
 */
function getSongsByFirstLetter(letter) {
  loadIndex();
  const lower = letter.toLowerCase();
  const results = [];

  for (const artist in cachedSongsByArtist) {
    if (artist.toLowerCase().startsWith(lower)) {
      cachedSongsByArtist[artist].forEach(song => results.push(song));
    }
  }

  // Sort by artist, then title
  results.sort((a, b) => {
    const artistCompare = a.artist.toLowerCase().localeCompare(b.artist.toLowerCase());
    if (artistCompare !== 0) return artistCompare;
    return a.title.toLowerCase().localeCompare(b.title.toLowerCase());
  });

  return results;
}




/**
 * Instant search using the indexed dictionary
 */
function searchSongs(term) {
  loadIndex();
  const lower = term.toLowerCase();

  const results = [];

  for (const artist in cachedSongsByArtist) {
    cachedSongsByArtist[artist].forEach(song => {
      if (
        song.artist.toLowerCase().includes(lower) ||
        song.title.toLowerCase().includes(lower)
      ) {
        results.push(song);
      }
    });
  }

  // ⭐ Sort results alphabetically
  results.sort((a, b) => {
    const artistCompare = a.artist.toLowerCase().localeCompare(b.artist.toLowerCase());
    if (artistCompare !== 0) return artistCompare;
    return a.title.toLowerCase().localeCompare(b.title.toLowerCase());
  });

  return results;
}


function getAllSongs() {
  loadIndex();

  const all = [];

  for (const artist in cachedSongsByArtist) {
    cachedSongsByArtist[artist].forEach(song => all.push(song));
  }

  // ⭐ Sort by artist, then title
  all.sort((a, b) => {
    const artistCompare = a.artist.toLowerCase().localeCompare(b.artist.toLowerCase());
    if (artistCompare !== 0) return artistCompare;
    return a.title.toLowerCase().localeCompare(b.title.toLowerCase());
  });

  return all;
}



/**
 * Submit song request
 */

function submitSongRequest(title, artist, name, key) {
    // Save key along with the request
  const ss = SpreadsheetApp.openById(SONGBOOK_SHEET_ID);
  const sheet = ss.getSheetByName(REQUEST_SHEET_NAME);

  // Force Google Sheets to treat key as text
key = "'" +key;

  sheet.appendRow([new Date(), name, artist, title, key]);
}

/**
 * Serve HTML
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .addMetaTag('viewport', 'width=device-width, initial-scale=0.9')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Optional: Helps if embedding elsewhere
}
