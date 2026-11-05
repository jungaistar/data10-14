/* 브라우저 저장소 — localStorage 를 쓰되, 막혀 있거나 용량이 차면 메모리로만 동작합니다.
   이 기기 사람 정보(작성자 이름·기기 ID)는 실사 자료와 따로 둡니다(백업·결과 파일에 섞이지 않게). */
(function (root) {
  'use strict';
  var KEY_DB = 'data10-14.db', KEY_ME = 'data10-14.me';
  var memory = {};
  var ok = true;
  var lastError = '';
  function get(k) {
    try { var v = root.localStorage.getItem(k); return v == null && memory[k] != null ? memory[k] : v; }
    catch (e) { ok = false; return memory[k] == null ? null : memory[k]; }
  }
  function set(k, v) {
    memory[k] = v;
    try { root.localStorage.setItem(k, v); lastError = ''; return true; }
    catch (e) {
      ok = false;
      lastError = (e && (e.name === 'QuotaExceededError' || e.code === 22)) ? 'quota' : 'blocked';
      return false;
    }
  }
  function del(k) {
    delete memory[k];
    try { root.localStorage.removeItem(k); } catch (e) { ok = false; }
  }
  root.STStore = {
    loadDb: function () {
      var raw = get(KEY_DB);
      if (!raw) return root.STLogic.emptyDb();
      try { return root.STLogic.restoreDb(JSON.parse(raw)); } catch (e) { return root.STLogic.emptyDb(); }
    },
    /* true = 브라우저에 저장됨, false = 이번 창의 메모리에만 있음 */
    saveDb: function (db) { return set(KEY_DB, JSON.stringify(db)); },
    clearDb: function () { del(KEY_DB); },
    loadMe: function () {
      var me = null;
      try { me = JSON.parse(get(KEY_ME) || 'null'); } catch (e) { me = null; }
      if (!me || !me.dev) me = { name: me && me.name || '', dev: 'd' + Math.random().toString(36).slice(2, 8), wh: '' };
      set(KEY_ME, JSON.stringify(me));
      return me;
    },
    saveMe: function (me) { return set(KEY_ME, JSON.stringify(me)); },
    lastError: function () { return lastError; },
    sizeKb: function () { var v = get(KEY_DB); return v ? Math.round(v.length * 2 / 1024) : 0; },
    available: function () { get(KEY_DB); return ok; }
  };
})(window);
