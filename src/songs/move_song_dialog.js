/**
 * The Move a Song dialog (src/songs/move_song.js), kept as a string in a .js file so it needs no
 * separate HTML file in the Apps Script project.
 *
 * It loads every song once (getMoveSongData), previews the move as you type - which song, where it
 * lands, which songs shift - and only then calls moveSongFromDialog, which checks everything again
 * on the server before changing anything.
 */
function moveSongDialogHtml() {
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
  body { font: 14px/1.45 Arial, sans-serif; color: #202124; margin: 0; padding: 4px 2px; }
  label { display: block; font-weight: bold; margin: 12px 0 4px; }
  input[type=number] { width: 90px; font-size: 15px; padding: 4px 6px; }
  input[type=search] { width: 100%; box-sizing: border-box; font-size: 14px; padding: 4px 6px; }
  select { font-size: 14px; padding: 3px; max-width: 100%; }
  .hint { color: #5f6368; margin-left: 8px; }
  .row { display: flex; align-items: baseline; flex-wrap: wrap; }
  .box { background: #f1f3f4; border-radius: 6px; padding: 8px 10px; margin-top: 12px; min-height: 20px; }
  .warn { color: #b06000; }
  .bad { color: #c5221f; }
  .ok { color: #188038; }
  #matches { max-height: 110px; overflow-y: auto; margin-top: 4px; }
  #matches div { cursor: pointer; padding: 1px 4px; }
  #matches div:hover { background: #e8f0fe; }
  .buttons { margin-top: 16px; display: flex; gap: 8px; }
  button { font-size: 14px; padding: 6px 16px; border-radius: 4px; border: 1px solid #dadce0; background: #fff; cursor: pointer; }
  button.primary { background: #1a73e8; color: #fff; border-color: #1a73e8; }
  button:disabled { opacity: .5; cursor: default; }
  pre { white-space: pre-wrap; font: 13px/1.4 Arial, sans-serif; margin: 0; }
</style>
</head>
<body>
<div id="form">
  <label for="find">Find a song <span class="hint">(optional - click a result to use its row)</span></label>
  <input type="search" id="find" placeholder="Part of a title">
  <div id="matches"></div>

  <label for="from">Move row</label>
  <div class="row"><input type="number" id="from"><span class="hint" id="fromInfo"></span></div>

  <label for="to">To row <span class="hint">(the row it ends up on)</span></label>
  <div class="row"><input type="number" id="to"><span class="hint" id="toInfo"></span></div>

  <label for="category">Category afterwards</label>
  <select id="category"></select>

  <div class="box" id="preview">Loading the Tracklist...</div>
  <div class="buttons">
    <button class="primary" id="move" disabled>Move</button>
    <button id="cancel">Cancel</button>
  </div>
</div>
<div id="result" style="display:none">
  <div class="box"><pre id="resultText"></pre></div>
  <div class="buttons">
    <button class="primary" id="another">Move another</button>
    <button id="close">Close</button>
  </div>
</div>

<script>
  var data = null, categoryTouched = false;
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };

  function load() {
    $('preview').textContent = 'Loading the Tracklist...';
    google.script.run.withSuccessHandler(function (d) {
      data = d;
      data.byRow = {};
      d.songs.forEach(function (s) { data.byRow[s.row] = s; });
      var sel = $('category');
      sel.innerHTML = '<option value="">(no category)</option>' +
        d.categories.map(function (c) { return '<option>' + esc(c) + '</option>'; }).join('');
      $('from').min = $('to').min = d.firstRow;
      $('from').max = $('to').max = d.lastRow;
      update();
    }).withFailureHandler(function (e) {
      $('preview').innerHTML = '<span class="bad">Could not read the Tracklist: ' + esc(e.message) + '</span>';
    }).getMoveSongData();
  }

  function describe(s) {
    return s ? '"' + esc(s.title) + '" · ' + esc(s.category || 'no category') + (s.status !== 'active' ? ' · ' + esc(s.status) : '') : '';
  }

  // Where each song would be after moving row "from" to row "to".
  function after(from, to) {
    return data.songs.map(function (s) {
      var row = s.row;
      if (row === from) row = to;
      else if (from < to && row > from && row <= to) row--;
      else if (to < from && row >= to && row < from) row++;
      return { row: row, song: s };
    });
  }

  function update() {
    if (!data) return;
    var from = parseInt($('from').value, 10), to = parseInt($('to').value, 10);
    var song = data.byRow[from];
    var problems = [], notes = [];
    $('fromInfo').innerHTML = isNaN(from) ? '' : song ? describe(song) : '<span class="bad">no song on this row</span>';

    var above = null, below = null;
    if (!isNaN(to)) {
      if (to < data.firstRow || to > data.lastRow) {
        problems.push('The songs are on rows ' + data.firstRow + '-' + data.lastRow + '.');
      } else if (song) {
        var placed = {};
        after(from, to).forEach(function (p) { placed[p.row] = p.song; });
        above = placed[to - 1] || null;
        below = placed[to + 1] || null;
      }
    }
    $('toInfo').innerHTML = song && !isNaN(to) && !problems.length
      ? 'between ' + (above ? describe(above) : '(nothing)') + '<br>and ' + (below ? describe(below) : '(nothing)')
      : '';

    // Suggest a category, until one is picked by hand: keep the song's own if it borders it,
    // otherwise take the one it lands in.
    if (song && !categoryTouched) {
      var own = song.category;
      var suggested = (above && above.category === own) || (below && below.category === own) ? own
        : above ? above.category : below ? below.category : own;
      $('category').value = suggested || '';
    }
    var category = $('category').value;

    if (!song) problems.push(isNaN(from) ? 'Enter the row of the song to move.' : 'Row ' + from + ' has no song.');
    if (isNaN(to)) problems.push('Enter the row it should end up on.');
    if (song && from === to && category === song.category) problems.push('That is where it already is.');

    if (song && !problems.length) {
      if (from !== to) {
        var lo = Math.min(from, to), hi = Math.max(from, to);
        notes.push(from < to ? 'Rows ' + (from + 1) + '-' + to + ' move up one.' : 'Rows ' + to + '-' + (from - 1) + ' move down one.');
      }
      if (category !== song.category) notes.push('Category: ' + esc(song.category || 'none') + ' → ' + esc(category || 'none') + '. Its history in the archives moves with it.');
      if (category && !(above && above.category === category) && !(below && below.category === category)) {
        var others = data.songs.filter(function (s) { return s.category === category && s !== song; }).length;
        if (others) notes.push('<span class="warn">It won\\'t be next to any other ' + esc(category) + ' song, so that category will be split.</span>');
      }
    }

    $('preview').innerHTML = problems.length
      ? problems.map(function (p) { return '<span class="bad">' + esc(p) + '</span>'; }).join('<br>')
      : notes.join('<br>') || 'Ready.';
    $('move').disabled = problems.length > 0;
  }

  function search() {
    var q = $('find').value.trim().toLowerCase();
    if (!data || q.length < 2) { $('matches').innerHTML = ''; return; }
    var hits = data.songs.filter(function (s) { return s.title.toLowerCase().indexOf(q) !== -1; }).slice(0, 30);
    $('matches').innerHTML = hits.length
      ? hits.map(function (s) { return '<div data-row="' + s.row + '">' + s.row + ' · ' + esc(s.title) + ' <span class="hint">' + esc(s.category) + '</span></div>'; }).join('')
      : '<span class="hint">No song with that in its title.</span>';
  }

  function move() {
    var from = parseInt($('from').value, 10), to = parseInt($('to').value, 10), category = $('category').value;
    $('move').disabled = true;
    $('cancel').disabled = true;
    $('preview').textContent = 'Moving... this rewrites every archive and can take a minute. Please don\\'t edit the spreadsheet.';
    google.script.run.withSuccessHandler(function (r) {
      $('form').style.display = 'none';
      $('result').style.display = 'block';
      $('resultText').className = r.ok ? '' : 'bad';
      $('resultText').textContent = r.message;
    }).withFailureHandler(function (e) {
      $('form').style.display = 'none';
      $('result').style.display = 'block';
      $('resultText').className = 'bad';
      $('resultText').textContent = 'Error: ' + e.message + '\\n\\nRun Checks → Check Row Alignment before anything else.';
    }).moveSongFromDialog(from, to, category);
  }

  $('from').addEventListener('input', function () { categoryTouched = false; update(); });
  $('to').addEventListener('input', update);
  $('category').addEventListener('change', function () { categoryTouched = true; update(); });
  $('find').addEventListener('input', search);
  $('matches').addEventListener('click', function (e) {
    var el = e.target.closest('[data-row]');
    if (!el) return;
    $('from').value = el.getAttribute('data-row');
    categoryTouched = false;
    update();
    $('to').focus();
  });
  $('move').addEventListener('click', move);
  $('cancel').addEventListener('click', function () { google.script.host.close(); });
  $('close').addEventListener('click', function () { google.script.host.close(); });
  $('another').addEventListener('click', function () {
    $('result').style.display = 'none';
    $('form').style.display = 'block';
    $('cancel').disabled = false;
    $('from').value = ''; $('to').value = '';
    categoryTouched = false;
    load();
  });
  load();
</script>
</body>
</html>`;
}
