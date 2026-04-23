// 日程情報（大学前教室 全15回）
let sessionsInfo = [
    { id: 'day1',  date: '4/21(火)', title: 'Day 1' },
    { id: 'day2',  date: '4/22(水)', title: 'Day 2' },
    { id: 'day3',  date: '4/23(木)', title: 'Day 3' },
    { id: 'day4',  date: '4/28(火)', title: 'Day 4' },
    { id: 'day5',  date: '4/29(水)', title: 'Day 5' },
    { id: 'day6',  date: '4/30(木)', title: 'Day 6' },
    { id: 'day7',  date: '5/12(火)', title: 'Day 7' },
    { id: 'day8',  date: '5/13(水)', title: 'Day 8' },
    { id: 'day9',  date: '5/14(木)', title: 'Day 9' },
    { id: 'day10', date: '5/19(火)', title: 'Day 10' },
    { id: 'day11', date: '5/20(水)', title: 'Day 11' },
    { id: 'day12', date: '5/21(木)', title: 'Day 12' },
    { id: 'day13', date: '5/26(火)', title: 'Day 13' },
    { id: 'day14', date: '5/27(水)', title: 'Day 14' },
    { id: 'day15', date: '5/28(木)', title: 'Day 15' },
];

// 参加者リスト（フォーム回答から自動取得）
let participantsList = [];

// 過去問の選択肢を生成する（古い順）
function getPastPaperOptions(grade) {
    let options = ['<option value="">未選択/その他</option>'];
    
    for (let year = 2018; year <= 2025; year++) {
        // 準2級プラスは2025年新設のためスキップ
        if (grade === '準2級プラス' && year < 2025) continue;
        
        for (let num = 1; num <= 3; num++) {
            let label = `第${num}回`;
            options.push(`<option value="${year}年度 ${label}">${year}年度 ${label}</option>`);
            
            let satLabel = `第${num}回（準会場）`;
            options.push(`<option value="${year}年度 ${satLabel}">${year}年度 ${satLabel}</option>`);
        }
    }
    
    return options.join('\n');
}

// ★ 受験級ごとのカラーマッピング（ライト/ダーク両対応）
function getGradeColor(grade) {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    const colors = isDark ? {
        '5級': '#60a5fa',         // 明るい青
        '4級': '#34d399',         // 明るい緑
        '3級': '#fbbf24',         // 明るいオレンジ
        '準2級': '#f472b6',       // 明るいピンク
        '準2級プラス': '#e879f9', // 明るいマゼンタ
        '2級': '#a78bfa',         // 明るい紫
        '準1級': '#f87171',       // 明るい赤
        '1級': '#fb7185',         // 明るいローズ
    } : {
        '5級': '#3b82f6',    // 青
        '4級': '#10b981',    // 緑
        '3級': '#f59e0b',    // オレンジ
        '準2級': '#e84393',  // ピンク
        '準2級プラス': '#d946ef', // マゼンタ
        '2級': '#7c3aed',    // 紫
        '準1級': '#dc2626',  // 深紅
        '1級': '#991b1b',    // ダーク赤
    };
    return colors[grade] || '#94a3b8';
}

// ====== GAS Web API URL ======
// ★★★ GASデプロイ後に取得したURLをここに貼り付けてください ★★★
const GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyFjUMbynTewZ4LR9jl4fUdtc8WNRh-_-UlhY771JZF-b9X_jAt4Ujah6l8cvozChE7aw/exec";

// アプリの状態
let currentSessionId = null;
let appData = {};
let newParticipantIds = new Set(); // 未確認の新規参加者を追跡（localStorageに永続化）
let deletedParticipantNames = []; // ★ 削除済み参加者名リスト（フォーム再処理で復活防止）
let hasRendered = false; // ★ 画面が一度でも描画されたかどうか（初回保存防止用）
let currentSort = { key: null, asc: true }; // ★ ソート状態
let cloudLastUpdated = null; // ★ クラウド最終更新時刻（デバッグ用）
let isSyncing = false; // ★ 同期中フラグ（同期中の保存を防止）

// 初期化
function init() {
    // 保存済みの日程を読み込む
    const savedSessions = localStorage.getItem('eikenDaigakumaeSessions');
    if (savedSessions) {
        sessionsInfo = JSON.parse(savedSessions);
    }
    // 未確認のNEWバッジを復元
    const savedNewIds = localStorage.getItem('eikenDaigakumaeNewIds');
    if (savedNewIds) {
        newParticipantIds = new Set(JSON.parse(savedNewIds));
    }
    const savedDeleted = localStorage.getItem('eikenDaigakumaeDeleted');
    if (savedDeleted) {
        deletedParticipantNames = JSON.parse(savedDeleted);
    }
    // テーマトグルのアイコンを初期化
    updateThemeIcon();
    loadData();
    renderSidebar();
}

// ====== テーマ切替（ライト / ダーク） ======
function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
    const next = current === 'dark' ? 'light' : 'dark';
    if (next === 'dark') {
        document.documentElement.setAttribute('data-theme', 'dark');
    } else {
        document.documentElement.removeAttribute('data-theme');
    }
    try { localStorage.setItem('eikenTheme', next); } catch (e) {}
    updateThemeIcon();
    // ★ 受験級の色分けもテーマに合わせて再描画
    if (currentSessionId && hasRendered) {
        renderMainContent();
    }
}

function updateThemeIcon() {
    const btn = document.getElementById('themeToggleBtn');
    if (!btn) return;
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    btn.textContent = isDark ? '☀️' : '🌙';
    btn.title = isDark ? 'ライトモードに切り替え' : 'ダークモードに切り替え';
}

// アプリ初期化・データ同期
async function loadData() {
    // 1. オフラインのバックアップをとりあえず読み込む
    const savedP = localStorage.getItem('eikenDaigakumaeParticipants');
    if (savedP) {
        participantsList = JSON.parse(savedP);
    }

    const saved = localStorage.getItem('eikenDaigakumaeData');
    if (saved) {
        appData = JSON.parse(saved);
        sessionsInfo.forEach(session => {
            if (!appData[session.id]) {
                appData[session.id] = { generalHomework: '', generalNotes: '', participants: {} };
            }
        });
    } else {
        sessionsInfo.forEach(session => {
            appData[session.id] = { generalHomework: '', generalNotes: '', participants: {} };
        });
    }
    
    // 2. 起動時に自動でクラウド同期（フォーム回答＆成績）
    await syncWithCloud();
}

// ====== クラウド同期機能 ======
async function syncWithCloud() {
    if (isSyncing) return;
    isSyncing = true;
    showLoading("クラウドから最新のデータを同期中...");
    try {
        // キャッシュを回避するためクエリにタイムスタンプを付与
        const url = GAS_WEB_APP_URL + (GAS_WEB_APP_URL.includes('?') ? '&' : '?') + 't=' + Date.now();
        const response = await fetch(url, { cache: 'no-store' });
        const data = await response.json();
        cloudLastUpdated = data.lastUpdated || null;

        // ★ 1. クラウドの成績データをフル上書き（クラウド優先）
        if (data.appData && Object.keys(data.appData).length > 0) {
            appData = data.appData;
        }

        // ★ 2. クラウドの participantsList があれば採用（他端末での氏名・級・手動追加の編集を反映）
        if (Array.isArray(data.participantsList) && data.participantsList.length > 0) {
            participantsList = data.participantsList;
        }

        // ★ 3. クラウドの sessionsInfo があれば採用（他端末で追加された日程を反映）
        if (Array.isArray(data.sessionsInfo) && data.sessionsInfo.length > 0) {
            sessionsInfo = data.sessionsInfo;
            localStorage.setItem('eikenDaigakumaeSessions', JSON.stringify(sessionsInfo));
        }

        // ★ 3b. クラウドの削除済み参加者リストがあれば採用
        if (Array.isArray(data.deletedParticipantNames)) {
            deletedParticipantNames = data.deletedParticipantNames;
        }

        // 全セッションのエントリを保証
        sessionsInfo.forEach(session => {
            if (!appData[session.id]) {
                appData[session.id] = { generalHomework: '', generalNotes: '', participants: {} };
            }
        });

        // ★ 4. フォーム回答を処理（新規参加者を追加）
        newParticipantIds = new Set();
        if (data.formResponses && data.formResponses.length > 0) {
            processFormResponses(data.formResponses);
        }

        // ★ 4b. フォームから削除された参加者を除去（フォーム回答をSource of Truthとする）
        // フォーム回答が存在する場合のみ実行（空の場合は安全のためスキップ）
        if (data.formResponses && data.formResponses.length > 0) {
            // ★ 名前の比較は「前後空白除去＋全角/半角空白畳み込み」で正規化して行う
            const _normName = (s) => String(s || '').replace(/[\s　]+/g, ' ').trim();
            const formNames = new Set();
            data.formResponses.forEach(row => {
                const nameKey = Object.keys(row).find(k => k.includes('氏名'));
                if (nameKey) {
                    const name = _normName(row[nameKey]);
                    if (name) formNames.add(name);
                }
            });

            // p_form_ IDの参加者で、現在のフォーム回答に存在しない人を除去
            const removedNames = [];
            participantsList = participantsList.filter(p => {
                if (!p.id.startsWith('p_form_')) return true; // 手動追加(p_manual_)は保持
                if (formNames.has(_normName(p.name))) return true; // フォームに存在する人は保持
                // フォームから消えた人 → 全セッションから除去
                Object.keys(appData).forEach(sessionId => {
                    if (appData[sessionId] && appData[sessionId].participants && appData[sessionId].participants[p.id]) {
                        delete appData[sessionId].participants[p.id];
                    }
                });
                removedNames.push(p.name);
                return false; // participantsListからも除去
            });
            if (removedNames.length > 0) {
                console.log('[syncWithCloud] フォームから削除された参加者を除去:', removedNames);
            }
        }

        // ★ 5. 旧IDのゴミデータをクリーンアップ（participantsListに存在しないIDを除去）
        const validIds = new Set(participantsList.map(p => p.id));
        Object.keys(appData).forEach(sessionId => {
            if (appData[sessionId] && appData[sessionId].participants) {
                Object.keys(appData[sessionId].participants).forEach(pid => {
                    if (!validIds.has(pid)) {
                        delete appData[sessionId].participants[pid];
                    }
                });
            }
        });

        // 新規参加者がいれば通知を表示
        if (newParticipantIds.size > 0) {
            showNotification(`🆕 新しい申し込みが ${newParticipantIds.size} 件あります！`);
            localStorage.setItem('eikenDaigakumaeNewIds', JSON.stringify([...newParticipantIds]));
        }

        // ★ ローカルにも最新データを保存
        localStorage.setItem('eikenDaigakumaeData', JSON.stringify(appData));
        localStorage.setItem('eikenDaigakumaeParticipants', JSON.stringify(participantsList));
        localStorage.setItem('eikenDaigakumaeDeleted', JSON.stringify(deletedParticipantNames));

        if (!currentSessionId && sessionsInfo.length > 0) {
            currentSessionId = sessionsInfo[0].id;
        }
        renderSidebar();
        if (currentSessionId) {
            renderMainContent();
            updateHeader();
        }
    } catch(e) {
        console.error("クラウド同期エラー:", e);
    } finally {
        isSyncing = false;
        hideLoading();
    }
}

function updateHeader() {
    if (!currentSessionId) return;
    const sessionInfo = sessionsInfo.find(s => s.id === currentSessionId);
    if (sessionInfo) {
        document.getElementById('currentDateTitle').textContent = sessionInfo.date;
        document.getElementById('currentSessionInfo').textContent = '';
        document.getElementById('emptyState').style.display = 'none';
        document.getElementById('mainScrollable').style.display = 'block';
    }
}

function showLoading(msg) {
    const overlay = document.getElementById('loadingOverlay');
    const text = document.getElementById('loadingText');
    if (overlay && text) {
        text.textContent = msg || "通信中...";
        overlay.classList.add('active');
    }
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.classList.remove('active');
    }
}

// ★ 名前からデバイス非依存の決定的IDを生成
function generateParticipantId(name, index) {
    let hash = 0;
    const str = name + '_' + index;
    for (let i = 0; i < str.length; i++) {
        const char = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // Convert to 32bit integer
    }
    return 'p_form_' + Math.abs(hash).toString(36);
}

function processFormResponses(formResponses) {
    // ★ 名前の正規化（前後空白除去＋全角/半角空白の畳み込み）
    const _normName = (s) => String(s || '').replace(/[\s　]+/g, ' ').trim();

    // ★ 同じ人が複数行にわたって申し込んでいる場合に、
    //   「行Aで day1 を追加 → 行Bで day1 を削除」と打ち消し合うのを防ぐため、
    //   まず名前単位で参加希望を集約する。
    const aggregated = new Map(); // key: 正規化済みの名前 → { firstIndex, schoolYear, grade, rawAttendParts[], hasExplicitAttend }

    formResponses.forEach((row, i) => {
        const keys = Object.keys(row);
        let nameKey = keys.find(k => k.includes('氏名'));
        let yearKey = keys.find(k => k.includes('学年'));
        let gradeKey = keys.find(k => k.includes('級'));
        let attendKey = keys.find(k => k.includes('参加したい日') || k.includes('参加'));

        if (!nameKey) return;

        const rawName = String(row[nameKey] || '');
        if (!rawName.trim()) return;
        const name = _normName(rawName);

        // ★ 削除済み参加者はスキップ（フォーム回答からの復活を防止）
        if (deletedParticipantNames.includes(name)) return;

        const schoolYear = yearKey ? String(row[yearKey] || '') : '';
        const gradeStr = gradeKey ? String(row[gradeKey]) : '';
        let grade = '未選択';
        if (gradeStr.includes('5級')) grade = '5級';
        else if (gradeStr.includes('4級')) grade = '4級';
        else if (gradeStr.includes('3級')) grade = '3級';
        else if (gradeStr.includes('準2級プラス') || gradeStr.includes('準2級+')) grade = '準2級プラス';
        else if (gradeStr.includes('準2級')) grade = '準2級';
        else if (gradeStr.includes('2級')) grade = '2級';
        else if (gradeStr.includes('準1級')) grade = '準1級';
        else if (gradeStr.includes('1級')) grade = '1級';

        const rawAttend = attendKey ? String(row[attendKey] || '') : '';

        if (!aggregated.has(name)) {
            aggregated.set(name, {
                firstIndex: i,
                schoolYear: schoolYear,
                grade: grade,
                rawAttendParts: [],
                hasExplicitAttend: false
            });
        }
        const agg = aggregated.get(name);
        // 学年・級は最初の行の値を優先（空のときは後の行で補完）
        if (!agg.schoolYear && schoolYear) agg.schoolYear = schoolYear;
        if ((!agg.grade || agg.grade === '未選択') && grade !== '未選択') agg.grade = grade;
        if (rawAttend) {
            agg.rawAttendParts.push(rawAttend);
            agg.hasExplicitAttend = true;
        }
    });

    // 集約済みの参加希望を基に、participantsList と appData を更新
    aggregated.forEach((agg, name) => {
        // ★ 既存検索も正規化したうえで比較（古いデータに空白が混じっていても一致させる）
        let existing = participantsList.find(p => _normName(p.name) === name);
        if (!existing) {
            existing = {
                id: generateParticipantId(name, agg.firstIndex),
                name: name,
                grade: agg.grade,
                hasTablet: true, // ★ 大学前教室はタブレット持参が前提
                schoolYear: agg.schoolYear
            };
            participantsList.push(existing);
            newParticipantIds.add(existing.id);
        } else {
            // 既存ユーザーの手動編集は上書きしない（未入力のときだけフォーム値で補完）
            if (!existing.grade || existing.grade === '未選択') existing.grade = agg.grade;
            if (existing.hasTablet === undefined) existing.hasTablet = true;
            if (!existing.schoolYear) existing.schoolYear = agg.schoolYear;
            // ★ 旧データで空白混じりで保存されていた名前をこの機会に正規化
            if (existing.name !== name) existing.name = name;
        }

        const combinedRawAttend = agg.rawAttendParts.join(' , ');

        // 全日程にわたって参加希望を反映
        Object.keys(appData).forEach(sessionId => {
            const sessionData = sessionsInfo.find(s => s.id === sessionId);
            if (!sessionData) return;

            // ★ この日程から個別に除外された参加者はスキップ
            const excludedIds = appData[sessionId].excludedParticipantIds || [];
            if (excludedIds.includes(existing.id)) return;

            if (agg.hasExplicitAttend) {
                const dateStr = sessionData.date.substring(0, sessionData.date.indexOf('(') > -1 ? sessionData.date.indexOf('(') : sessionData.date.length);
                const wantsToAttend = combinedRawAttend.includes(dateStr);

                if (wantsToAttend) {
                    if (!appData[sessionId].participants[existing.id]) {
                        appData[sessionId].participants[existing.id] = {
                            attended: true, rpContent: '', rpScore: '', apContent: '', apScore: '', remarks: ''
                        };
                    }
                } else if (appData[sessionId].participants[existing.id]) {
                    const pData = appData[sessionId].participants[existing.id];
                    // スコア・備考が未入力なら削除（フォームでの希望変更を反映）
                    if (!pData.rpScore && !pData.apScore && !pData.remarks) {
                        delete appData[sessionId].participants[existing.id];
                    }
                }
            } else {
                // 参加希望日の指定がない場合：全日程に追加
                if (!appData[sessionId].participants[existing.id]) {
                    appData[sessionId].participants[existing.id] = {
                        attended: true, rpContent: '', rpScore: '', apContent: '', apScore: '', remarks: ''
                    };
                }
            }
        });
    });
}

// データの保存（クラウド＆ローカル保存）
// ★ フォーム+iframe送信方式でCORSを完全回避
async function saveData() {
    if (!currentSessionId) return;
    // ★ 同期中は保存しない（同期中にDOMから上書きされると整合性が崩れるため）
    if (isSyncing) return;
    // ★ まだ画面が描画されていない段階ではDOMから空データを取得してしまうため保存しない
    if (!hasRendered) return;

    appData[currentSessionId].generalHomework = document.getElementById('generalHomework').value;
    appData[currentSessionId].generalNotes = document.getElementById('generalNotes').value;

    if (appData[currentSessionId] && appData[currentSessionId].participants) {
        Object.keys(appData[currentSessionId].participants).forEach(id => {
            const attendElem = document.getElementById(`attend_${id}`);
            if (attendElem) {
                const attended = attendElem.checked;
                const rpContent = document.getElementById(`rpContent_${id}`).value;
                const rpScore = document.getElementById(`rpScore_${id}`).value;
                const apContent = document.getElementById(`apContent_${id}`).value;
                const apScore = document.getElementById(`apScore_${id}`).value;
                const remarks = document.getElementById(`remarks_${id}`).value;

                appData[currentSessionId].participants[id] = {
                    attended, rpContent, rpScore, apContent, apScore, remarks
                };
            }
        });
    }

    // オフライン動作用の一時保存
    localStorage.setItem('eikenDaigakumaeData', JSON.stringify(appData));
    localStorage.setItem('eikenDaigakumaeParticipants', JSON.stringify(participantsList));
    localStorage.setItem('eikenDaigakumaeSessions', JSON.stringify(sessionsInfo));
    localStorage.setItem('eikenDaigakumaeDeleted', JSON.stringify(deletedParticipantNames));

    // 保存メッセージの表示
    const status = document.getElementById('saveStatus');
    status.textContent = 'クラウドへ保存中...';
    status.classList.add('show');
    status.style.color = '';

    try {
        const result = await saveToCloud({
            appData: appData,
            participantsList: participantsList,
            sessionsInfo: sessionsInfo,
            deletedParticipantNames: deletedParticipantNames
        });
        if (result && result.status === 'success') {
            status.textContent = `クラウドへ保存完了 ✓ (参加者${result.savedParticipantsCount ?? '-'}名)`;
            status.style.color = '';
        } else {
            status.textContent = '保存されましたが検証未確認です';
            status.style.color = '#f59e0b';
        }
        setTimeout(() => { status.classList.remove('show'); }, 4000);
    } catch (e) {
        console.error("クラウド保存エラー:", e);
        status.textContent = '⚠ クラウド保存失敗: ' + (e.message || '不明');
        status.style.color = '#dc2626';
        setTimeout(() => { status.classList.remove('show'); }, 8000);
    }
}

// ★ 現在のメモリ状態のみをクラウドへ保存（DOMからは読み取らない）
// 参加者情報の編集・日程追加・手動追加・削除の直後に呼び出す
async function pushStateToCloud() {
    if (isSyncing) return;
    localStorage.setItem('eikenDaigakumaeData', JSON.stringify(appData));
    localStorage.setItem('eikenDaigakumaeParticipants', JSON.stringify(participantsList));
    localStorage.setItem('eikenDaigakumaeSessions', JSON.stringify(sessionsInfo));
    localStorage.setItem('eikenDaigakumaeDeleted', JSON.stringify(deletedParticipantNames));

    const status = document.getElementById('saveStatus');
    if (status) {
        status.textContent = 'クラウドへ保存中...';
        status.classList.add('show');
        status.style.color = '';
    }
    try {
        const result = await saveToCloud({
            appData: appData,
            participantsList: participantsList,
            sessionsInfo: sessionsInfo,
            deletedParticipantNames: deletedParticipantNames
        });
        if (status) {
            if (result && result.status === 'success') {
                status.textContent = `クラウドへ保存完了 ✓ (参加者${result.savedParticipantsCount ?? '-'}名)`;
                status.style.color = '';
            } else {
                status.textContent = '保存されましたが検証未確認です';
                status.style.color = '#f59e0b';
            }
            setTimeout(() => { status.classList.remove('show'); }, 3000);
        }
    } catch (e) {
        console.error("クラウド保存エラー:", e);
        if (status) {
            status.textContent = '⚠ クラウド保存失敗: ' + (e.message || '不明');
            status.style.color = '#dc2626';
            setTimeout(() => { status.classList.remove('show'); }, 8000);
        }
    }
}

// ★ 編集イベントの「連打」から cloud を守るためのデバウンス
let _pushDebounceTimer = null;
function pushStateToCloudDebounced(delay) {
    if (_pushDebounceTimer) clearTimeout(_pushDebounceTimer);
    _pushDebounceTimer = setTimeout(() => {
        _pushDebounceTimer = null;
        pushStateToCloud();
    }, delay || 800);
}

// ★ GASへのPOST保存（fetch text/plain方式 → 成功/失敗を明確に検知）
// GAS の Web App は text/plain Content-Type を受け付ける（CORSプリフライトを発生させないため）
// レスポンスは不透明リダイレクト経由で取得できる。成功/失敗は JSON で返ってくる
async function saveToCloud(payload) {
    const jsonStr = JSON.stringify(payload);
    console.log('[saveToCloud] 送信開始:', {
        sessions: (payload.sessionsInfo || []).length,
        participants: (payload.participantsList || []).length,
        appDataKeys: Object.keys(payload.appData || {}).length,
        jsonSizeKB: (jsonStr.length / 1024).toFixed(1)
    });
    let response, text;
    try {
        response = await fetch(GAS_WEB_APP_URL, {
            method: 'POST',
            mode: 'cors',
            redirect: 'follow',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: jsonStr
        });
        text = await response.text();
    } catch (netErr) {
        console.error('[saveToCloud] ネットワークエラー:', netErr);
        // fetch が CORS 等で失敗した場合、iframe フォールバックを試す
        try {
            await _saveToCloudViaIframe(payload);
            console.warn('[saveToCloud] iframeフォールバックで送信完了（結果検証は GET 同期時に行います）');
            return { status: 'success-unverified', via: 'iframe' };
        } catch (fallbackErr) {
            throw new Error('クラウドへ接続できません: ' + netErr.message);
        }
    }

    let result;
    try { result = JSON.parse(text); } catch (_) { result = { status: 'error', message: '無効なレスポンス: ' + (text || '').substring(0, 200) }; }
    console.log('[saveToCloud] レスポンス:', result);

    if (result.status !== 'success') {
        throw new Error('クラウド保存失敗: ' + (result.message || 'unknown'));
    }
    return result;
}

// フォールバック: 隠しiframe経由のPOST（fetchでCORSに弾かれた場合用）
function _saveToCloudViaIframe(payload) {
    return new Promise((resolve) => {
        const oldIframe = document.getElementById('gas_save_frame');
        if (oldIframe) oldIframe.remove();
        const oldForm = document.getElementById('gas_save_form');
        if (oldForm) oldForm.remove();

        const iframe = document.createElement('iframe');
        iframe.id = 'gas_save_frame';
        iframe.name = 'gas_save_frame';
        iframe.style.display = 'none';
        document.body.appendChild(iframe);

        const form = document.createElement('form');
        form.id = 'gas_save_form';
        form.method = 'POST';
        form.action = GAS_WEB_APP_URL;
        form.target = 'gas_save_frame';

        const jsonStr = JSON.stringify(payload);
        const utf8Bytes = new TextEncoder().encode(jsonStr);
        let binaryStr = '';
        utf8Bytes.forEach(b => binaryStr += String.fromCharCode(b));
        const base64Data = btoa(binaryStr);

        const input = document.createElement('input');
        input.type = 'hidden';
        input.name = 'data_b64';
        input.value = base64Data;
        form.appendChild(input);

        document.body.appendChild(form);
        form.submit();

        setTimeout(() => {
            iframe.remove();
            form.remove();
            resolve();
        }, 3000);
    });
}

// サイドバーの描画
function renderSidebar() {
    const list = document.getElementById('dateList');
    list.innerHTML = '';

    sessionsInfo.forEach(session => {
        const li = document.createElement('li');
        li.onclick = () => selectSession(session.id);
        
        // アクティブ状態の保持
        if (currentSessionId === session.id) {
            li.className = 'active';
        }

        // ★ 各日程の参加人数を計算（participantsListに存在する生徒のみ）
        let count = 0;
        if (appData[session.id] && appData[session.id].participants) {
            count = Object.keys(appData[session.id].participants).filter(pid => 
                participantsList.some(p => p.id === pid)
            ).length;
        }

        li.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
                <span class="date-title">${session.date}</span>
                <span style="background: ${currentSessionId === session.id ? 'rgba(255,255,255,0.25)' : 'var(--input-bg)'}; color: ${currentSessionId === session.id ? '#fff' : 'var(--text-sec)'}; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 10px; min-width: 28px; text-align: center;">${count}名</span>
            </div>
        `;
        list.appendChild(li);
    });
}

// セッション（日程）の選択
async function selectSession(sessionId) {
    await saveData(); // 今の画面を保存してから遷移
    currentSessionId = sessionId;
    
    const sessionInfo = sessionsInfo.find(s => s.id === sessionId);
    document.getElementById('currentDateTitle').textContent = sessionInfo.date;
    document.getElementById('currentSessionInfo').textContent = '';

    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('mainScrollable').style.display = 'block';

    renderSidebar(); // アクティブハイライトの更新
    renderMainContent();
}

// サイドバーの開閉トグル
function toggleSidebar() {
    const container = document.querySelector('.app-container');
    container.classList.toggle('sidebar-hidden');
}

// 参加者情報の直接編集
function updateParticipantInfo(id, field, value) {
    const p = participantsList.find(x => x.id === id);
    if (p) {
        if (field === 'hasTablet') {
            p[field] = value === 'true';
        } else {
            p[field] = value;
        }
        localStorage.setItem('eikenDaigakumaeParticipants', JSON.stringify(participantsList));

        // 受験級が変更された場合はプルダウンを再描画する
        if (field === 'grade') {
            saveData();
            renderMainContent();
        }

        // ★ 参加者情報の編集もクラウドへ同期（デバウンス）
        pushStateToCloudDebounced(1000);
    }
}

// 当日飛び入り参加者の手動追加
function addParticipant() {
    if (!currentSessionId) {
        alert('参加者を追加したい日程をサイドバーから選択してください。');
        return;
    }
    
    saveData(); // 現在の内容を一時保存
    
    const newId = 'p_manual_' + Date.now();
    participantsList.push({
        id: newId,
        name: '新規参加',
        grade: '3級',
        hasTablet: true,
        schoolYear: ''
    });
    
    // 現在選択中の日程だけに参加者を追加
    if (!appData[currentSessionId]) {
        appData[currentSessionId] = { generalHomework: '', generalNotes: '', participants: {} };
    }
    appData[currentSessionId].participants[newId] = {
        attended: true,
        rpContent: '',
        rpScore: '',
        apContent: '',
        apScore: '',
        remarks: ''
    };
    
    localStorage.setItem('eikenDaigakumaeParticipants', JSON.stringify(participantsList));
    localStorage.setItem('eikenDaigakumaeData', JSON.stringify(appData));

    renderMainContent();

    // ★ 手動追加した参加者を即座にクラウドへ同期（他端末で消えないように）
    pushStateToCloud();

    // 追加した参加者の名前入力欄にフォーカスを当てる（少し遅延させる）
    setTimeout(() => {
        const trs = document.getElementById('participantTableBody').querySelectorAll('tr');
        if (trs.length > 0) {
            const lastTr = trs[trs.length - 1];
            const nameInput = lastTr.querySelector('input[type="text"]');
            if (nameInput) nameInput.select();
        }
    }, 100);
}

// モーダル管理用
let participantToDelete = null;

// 参加者の削除（モーダル表示：この日 or 全日程の選択）
function deleteParticipant(id) {
    const p = participantsList.find(x => x.id === id);
    if (!p) return;
    
    participantToDelete = id;
    document.getElementById('deleteModalMessage').innerHTML = `<strong>${p.name}</strong> さんをどの範囲で削除しますか？`;
    document.getElementById('deleteModal').style.display = 'flex';
}

// モーダル閉じる
function closeDeleteModal() {
    document.getElementById('deleteModal').style.display = 'none';
    participantToDelete = null;
}

// この日のみ削除
function deleteFromCurrentSession() {
    if (!participantToDelete || !currentSessionId) return;
    const id = participantToDelete;

    // 現在の日程からのみ削除
    if (appData[currentSessionId] && appData[currentSessionId].participants[id]) {
        delete appData[currentSessionId].participants[id];
    }

    // ★ この日程の除外リストに追加（フォーム再処理で復活を防止）
    if (!appData[currentSessionId].excludedParticipantIds) {
        appData[currentSessionId].excludedParticipantIds = [];
    }
    if (!appData[currentSessionId].excludedParticipantIds.includes(id)) {
        appData[currentSessionId].excludedParticipantIds.push(id);
    }

    localStorage.setItem('eikenDaigakumaeData', JSON.stringify(appData));
    closeDeleteModal();
    renderMainContent();
    // ★ クラウドへも同期
    pushStateToCloud();
}

// 全日程から削除
// ★ 仕様: 現在の状態（participantsList と全日程のappData）からこの人を消すだけ。
//        deletedParticipantNames には追加しない。
//        フォーム回答にまだ行が残っていれば、次回同期時に空のスコアで再登場する。
//        完全に消したい場合は別途スプシの行も削除する必要がある。
function deleteFromAllSessions() {
    if (!participantToDelete) return;
    const id = participantToDelete;

    // 参加者リストから削除
    participantsList = participantsList.filter(x => x.id !== id);

    // 全日程のデータから削除
    Object.keys(appData).forEach(sessionId => {
        if (appData[sessionId].participants[id]) {
            delete appData[sessionId].participants[id];
        }
    });

    localStorage.setItem('eikenDaigakumaeParticipants', JSON.stringify(participantsList));
    localStorage.setItem('eikenDaigakumaeData', JSON.stringify(appData));
    closeDeleteModal();
    renderMainContent();
    // ★ クラウドへも同期
    pushStateToCloud();
}

// メインコンテンツ（参加者リストやフォーム）の描画
function renderMainContent() {
    const data = appData[currentSessionId];
    
    // 全体記録の復元
    document.getElementById('generalHomework').value = data.generalHomework || '';
    document.getElementById('generalNotes').value = data.generalNotes || '';

    // 参加者テーブルの描画
    const tbody = document.getElementById('participantTableBody');
    tbody.innerHTML = '';
    
    let renderedCount = 0;

    // ★ ソート済みリストを作成
    const gradeOrder = {'1級':1, '準1級':2, '2級':3, '準2級プラス':4, '準2級':5, '3級':6, '4級':7, '5級':8};
    let sortedList = participantsList.filter(p => data.participants[p.id]);
    
    if (currentSort.key) {
        sortedList.sort((a, b) => {
            let va, vb;
            if (currentSort.key === 'grade') {
                va = gradeOrder[a.grade] || 99;
                vb = gradeOrder[b.grade] || 99;
            } else if (currentSort.key === 'hasTablet') {
                va = a.hasTablet ? 0 : 1;
                vb = b.hasTablet ? 0 : 1;
            } else {
                va = (a[currentSort.key] || '').toString();
                vb = (b[currentSort.key] || '').toString();
            }
            if (va < vb) return currentSort.asc ? -1 : 1;
            if (va > vb) return currentSort.asc ? 1 : -1;
            return 0;
        });
    }
    
    // ソートインジケータ更新
    ['name', 'schoolYear', 'grade', 'hasTablet'].forEach(key => {
        const el = document.getElementById('sort_' + key);
        if (el) el.textContent = currentSort.key === key ? (currentSort.asc ? '▲' : '▼') : '';
    });

    sortedList.forEach(p => {
        const pData = data.participants[p.id];
        
        renderedCount++;
        const tr = document.createElement('tr');
        
        tr.innerHTML = `
            <td>
                <label class="toggle-switch">
                    <input type="checkbox" id="attend_${p.id}" ${pData.attended ? 'checked' : ''} onchange="toggleAttendance('${p.id}')">
                    <span class="slider"></span>
                </label>
            </td>
            <td>
                <input type="text" value="${p.name}" onchange="updateParticipantInfo('${p.id}', 'name', this.value)" style="width: 90px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px; font-weight: bold;">
                <button type="button" onclick="showSchedulePopup('${p.id}')" style="background:none;border:none;cursor:pointer;font-size:14px;padding:2px;vertical-align:middle;" title="参加日程を確認">📅</button>
                ${newParticipantIds.has(p.id) ? '<span class="badge-new">NEW</span>' : ''}
            </td>
            <td>
                <input type="text" value="${p.schoolYear || ''}" onchange="updateParticipantInfo('${p.id}', 'schoolYear', this.value)" style="width: 55px; text-align: center; border: 1px solid var(--border-color); border-radius: 4px; padding: 4px; font-size: 0.9em;">
            </td>
            <td>
                <select onchange="updateParticipantInfo('${p.id}', 'grade', this.value)" style="width: 75px; padding: 4px; border: 1.5px solid ${getGradeColor(p.grade)}; border-radius: 4px; background-color: ${getGradeColor(p.grade)}33; color: ${getGradeColor(p.grade)}; font-weight: 700; appearance: none; -webkit-appearance: none; -moz-appearance: none;">
                    ${['1級', '準1級', '2級', '準2級プラス', '準2級', '3級', '4級', '5級'].map(g => `<option value="${g}" ${p.grade === g ? 'selected' : ''}>${g}</option>`).join('')}
                </select>
            </td>
            <td>
                <select onchange="updateParticipantInfo('${p.id}', 'hasTablet', this.value)" style="width: 60px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px; color: ${p.hasTablet ? 'var(--primary-color)' : 'var(--text-color)'}; font-weight: ${p.hasTablet ? 'bold' : 'normal'};">
                    <option value="true" ${p.hasTablet ? 'selected' : ''}>持参</option>
                    <option value="false" ${!p.hasTablet ? 'selected' : ''}>なし</option>
                </select>
            </td>
            <td>
                <select id="rpContent_${p.id}" ${!pData.attended ? 'disabled' : ''} style="width: 210px; font-size: 13px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px;">
                    ${getPastPaperOptions(p.grade)}
                </select>
            </td>
            <td>
                <select id="rpScore_${p.id}" ${!pData.attended ? 'disabled' : ''} style="width: 60px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px; text-align: center;">
                    <option value="" ${!pData.rpScore ? 'selected' : ''}>-</option>
                    <option value="5" ${pData.rpScore === '5' ? 'selected' : ''}>5</option>
                    <option value="4" ${pData.rpScore === '4' ? 'selected' : ''}>4</option>
                    <option value="3" ${pData.rpScore === '3' ? 'selected' : ''}>3</option>
                    <option value="2" ${pData.rpScore === '2' ? 'selected' : ''}>2</option>
                    <option value="1" ${pData.rpScore === '1' ? 'selected' : ''}>1</option>
                </select>
            </td>
            <td>
                <select id="apContent_${p.id}" ${!pData.attended ? 'disabled' : ''} style="width: 210px; font-size: 13px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px;">
                    ${getPastPaperOptions(p.grade)}
                </select>
            </td>
            <td>
                <select id="apScore_${p.id}" ${!pData.attended ? 'disabled' : ''} style="width: 60px; padding: 4px; border: 1px solid var(--border-color); border-radius: 4px; text-align: center;">
                    <option value="" ${!pData.apScore ? 'selected' : ''}>-</option>
                    <option value="5" ${pData.apScore === '5' ? 'selected' : ''}>5</option>
                    <option value="4" ${pData.apScore === '4' ? 'selected' : ''}>4</option>
                    <option value="3" ${pData.apScore === '3' ? 'selected' : ''}>3</option>
                    <option value="2" ${pData.apScore === '2' ? 'selected' : ''}>2</option>
                    <option value="1" ${pData.apScore === '1' ? 'selected' : ''}>1</option>
                </select>
            </td>
            <td>
                <input type="text" id="remarks_${p.id}" value="${pData.remarks || ''}" placeholder="特記事項・個別宿題" ${!pData.attended ? 'disabled' : ''}>
            </td>
            <td>
                <button type="button" onclick="deleteParticipant('${p.id}')" style="background: none; border: none; font-size: 20px; cursor: pointer; opacity: 0.7; padding: 8px;" title="この参加者を削除">🗑️</button>
            </td>
        `;
        
        if (!pData.attended) {
            tr.style.opacity = '0.5';
        }
        
        tbody.appendChild(tr);
        
        // ★ innerHTML生成後にDOMで直接valueをセット（日本語文字列でも確実に動作）
        const rpSel = document.getElementById(`rpContent_${p.id}`);
        if (rpSel && pData.rpContent) rpSel.value = pData.rpContent;
        const apSel = document.getElementById(`apContent_${p.id}`);
        if (apSel && pData.apContent) apSel.value = pData.apContent;
    });
    
    document.getElementById('participantCount').textContent = `${renderedCount}名`;
    hasRendered = true; // ★ 描画完了フラグ
}

// 欠席時の入力制限
function toggleAttendance(id) {
    const isChecked = document.getElementById(`attend_${id}`).checked;
    
    // UIの切り替え
    const tr = document.getElementById(`attend_${id}`).closest('tr');
    if (isChecked) {
        tr.style.opacity = '1';
    } else {
        tr.style.opacity = '0.5';
    }
    
    document.getElementById(`rpContent_${id}`).disabled = !isChecked;
    document.getElementById(`rpScore_${id}`).disabled = !isChecked;
    document.getElementById(`apContent_${id}`).disabled = !isChecked;
    document.getElementById(`apScore_${id}`).disabled = !isChecked;
    document.getElementById(`remarks_${id}`).disabled = !isChecked;
}
// ====== 通知バナー ======
function showNotification(message) {
    const banner = document.getElementById('notificationBanner');
    const msgEl = document.getElementById('notifMessage');
    if (banner && msgEl) {
        msgEl.textContent = message;
        banner.classList.add('show');
    }
}

function dismissNotification() {
    const banner = document.getElementById('notificationBanner');
    if (banner) {
        banner.classList.remove('show');
    }
    // Newバッジもクリア（永続化も解除）
    newParticipantIds.clear();
    localStorage.removeItem('eikenDaigakumaeNewIds');
    if (currentSessionId) {
        renderMainContent();
    }
}

// ====== 参加日程ポップアップ ======
function showSchedulePopup(participantId) {
    const p = participantsList.find(x => x.id === participantId);
    if (!p) return;
    
    document.getElementById('schedulePopupName').textContent = p.name;
    
    const listEl = document.getElementById('schedulePopupList');
    listEl.innerHTML = '';
    
    // 今日の日付を取得（月日で比較用）
    const today = new Date();
    const todayStr = `${today.getMonth() + 1}月${today.getDate()}日`;
    
    sessionsInfo.forEach(session => {
        const isRegistered = appData[session.id] && appData[session.id].participants[participantId];
        
        // 日付文字列から月日を抽出して過去・未来を判定
        const dateMatch = session.date.match(/(\d+)月(\d+)日/);
        let isPast = false;
        if (dateMatch) {
            const sessionDate = new Date(today.getFullYear(), parseInt(dateMatch[1]) - 1, parseInt(dateMatch[2]));
            const todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
            isPast = sessionDate < todayMidnight;
        }
        
        const li = document.createElement('div');
        li.className = 'schedule-item' + (isRegistered ? ' registered' : ' not-registered') + (isPast ? ' past' : '');
        li.innerHTML = `
            <span class="schedule-status">${isRegistered ? '✅' : '—'}</span>
            <span class="schedule-date">${session.date}</span>
            ${isPast ? '<span class="schedule-past-label">済</span>' : ''}
        `;
        listEl.appendChild(li);
    });
    
    document.getElementById('schedulePopup').style.display = 'flex';
}

function closeSchedulePopup() {
    document.getElementById('schedulePopup').style.display = 'none';
}

// ====== 日程追加 ======
function showAddSessionModal() {
    document.getElementById('newSessionDate').value = '';
    document.getElementById('addSessionModal').style.display = 'flex';
}

function closeAddSessionModal() {
    document.getElementById('addSessionModal').style.display = 'none';
}

function addSession() {
    const dateValue = document.getElementById('newSessionDate').value;
    
    if (!dateValue) {
        alert('日付を選択してください。');
        return;
    }
    
    // yyyy-mm-dd → "X月Y日(曜日)" に変換
    const d = new Date(dateValue + 'T00:00:00');
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    const dateStr = `${d.getMonth() + 1}月${d.getDate()}日(${dayNames[d.getDay()]})`;
    
    const newId = 'day_' + Date.now();
    const newSession = {
        id: newId,
        date: dateStr,
        title: ''
    };
    
    sessionsInfo.push(newSession);
    appData[newId] = { generalHomework: '', generalNotes: '', participants: {} };

    // 保存
    localStorage.setItem('eikenDaigakumaeSessions', JSON.stringify(sessionsInfo));
    localStorage.setItem('eikenDaigakumaeData', JSON.stringify(appData));

    closeAddSessionModal();
    renderSidebar();

    // ★ 新しい日程を即座にクラウドへ同期（他端末にも反映）
    pushStateToCloud();
}

// ====== 空席状況・開講一覧モーダル ======
const SESSION_CAPACITY = 5;    // 各日程の定員
const MIN_TO_OPEN = 3;         // 開講に必要な最少人数

function getSessionCount(sessionId) {
    if (!appData[sessionId] || !appData[sessionId].participants) return 0;
    return Object.keys(appData[sessionId].participants).filter(pid =>
        participantsList.some(p => p.id === pid)
    ).length;
}

function showAvailabilityModal() {
    const listEl = document.getElementById('availabilityList');
    listEl.innerHTML = '';
    // コピーフィードバックをリセット
    const fb = document.getElementById('availCopyFeedback');
    if (fb) { fb.textContent = ''; fb.classList.remove('show'); }

    sessionsInfo.forEach(session => {
        const count = getSessionCount(session.id);
        const remaining = Math.max(0, SESSION_CAPACITY - count);
        const pct = Math.min((count / SESSION_CAPACITY) * 100, 100);

        let statusClass, statusLabel;
        if (count >= SESSION_CAPACITY) {
            statusClass = 'status-full';
            statusLabel = '満席';
        } else if (count < MIN_TO_OPEN) {
            statusClass = 'status-not-enough';
            statusLabel = `開講未定（あと${MIN_TO_OPEN - count}名）`;
        } else {
            statusClass = 'status-open';
            statusLabel = `開講予定（残${remaining}席）`;
        }

        const item = document.createElement('div');
        item.className = `avail-item ${statusClass}`;
        item.innerHTML = `
            <span class="avail-date">${session.date}</span>
            <div class="avail-bar-wrap">
                <div class="avail-bar">
                    <div class="avail-bar-fill" style="width: ${pct}%;"></div>
                </div>
                <span class="avail-count">${count}/${SESSION_CAPACITY}</span>
            </div>
            <span class="avail-status">${statusLabel}</span>
        `;
        listEl.appendChild(item);
    });

    // デフォルトの冒頭メッセージをセット（日付・時刻は自動）
    const msgEl = document.getElementById('availHeaderMsg');
    if (msgEl && !msgEl.value.trim()) {
        const now = new Date();
        const dateStr = `${now.getMonth() + 1}/${now.getDate()}付　${now.getHours()}:${String(now.getMinutes()).padStart(2, '0')}時点`;
        msgEl.value = `${dateStr}\n\nECCジュニア大学前教室の皆様へ\n英検勉強会の空き🈳状況です。\n当日キャンセルの方が出たため🈳が増えています。また人数減のため催行できない日程がある可能性があります。`;
    }

    document.getElementById('availabilityModal').style.display = 'flex';
}

function closeAvailabilityModal() {
    document.getElementById('availabilityModal').style.display = 'none';
}

// ★ LINE用テキスト生成＆コピー
function copyAvailabilityText() {
    const lines = [];

    // 冒頭メッセージを挿入
    const headerMsg = (document.getElementById('availHeaderMsg').value || '').trim();
    if (headerMsg) {
        lines.push(headerMsg);
        lines.push('ーーーーーー');
    }

    lines.push('📊 英検勉強会（大学前）空席状況');
    lines.push(`各日 定員${SESSION_CAPACITY}名／${MIN_TO_OPEN}名未満は開講しません`);
    lines.push('');

    sessionsInfo.forEach(session => {
        const count = getSessionCount(session.id);
        const remaining = Math.max(0, SESSION_CAPACITY - count);

        let icon, label;
        if (count >= SESSION_CAPACITY) {
            icon = '🈵';
            label = '満席';
        } else if (count < MIN_TO_OPEN) {
            icon = '⚠️';
            label = `残${remaining}席／開講未定（あと${MIN_TO_OPEN - count}名で開講）`;
        } else {
            icon = '⭕️';
            label = `残${remaining}席`;
        }
        lines.push(`${icon} ${session.date}　${label}`);
    });

    lines.push('');
    lines.push('⭕ …実施予定');
    lines.push('⚠️ …実施不可の可能性あり');
    lines.push('🈵 …キャンセル待ち');
    lines.push('');
    const now = new Date();
    lines.push(`（${now.getMonth() + 1}/${now.getDate()} ${now.getHours()}:${String(now.getMinutes()).padStart(2, '0')} 時点）`);

    const text = lines.join('\n');

    navigator.clipboard.writeText(text).then(() => {
        const fb = document.getElementById('availCopyFeedback');
        fb.textContent = '✓ コピーしました';
        fb.classList.add('show');
        setTimeout(() => fb.classList.remove('show'), 3000);
    }).catch(() => {
        // フォールバック（古いブラウザ対応）
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.cssText = 'position:fixed;left:-9999px;';
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        const fb = document.getElementById('availCopyFeedback');
        fb.textContent = '✓ コピーしました';
        fb.classList.add('show');
        setTimeout(() => fb.classList.remove('show'), 3000);
    });
}

// ====== テーブルソート ======
function sortTable(key) {
    if (currentSort.key === key) {
        currentSort.asc = !currentSort.asc; // 同じキー → 方向切り替え
    } else {
        currentSort.key = key;
        currentSort.asc = true;
    }
    renderMainContent();
}

// ====== フォーム回答スプレッドシートDL ======
async function downloadFormSpreadsheet() {
    showLoading("フォーム回答シートをダウンロード中...");
    try {
        const url = GAS_WEB_APP_URL + (GAS_WEB_APP_URL.includes('?') ? '&' : '?') + 'action=downloadFormSheet&t=' + Date.now();
        const response = await fetch(url, { cache: 'no-store' });
        const data = await response.json();

        if (data.status !== 'ok' || !data.rawData) {
            alert('エラー: ' + (data.message || 'データの取得に失敗しました'));
            return;
        }

        // GASから受け取った2次元配列をXLSXでExcel化
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(data.rawData);
        // 列幅を自動調整（各列の最大文字数を基準に設定）
        if (data.rawData.length > 0) {
            ws['!cols'] = data.rawData[0].map((_, colIdx) => {
                let maxLen = 8;
                data.rawData.forEach(row => {
                    const cellLen = String(row[colIdx] || '').length;
                    if (cellLen > maxLen) maxLen = cellLen;
                });
                return { wch: Math.min(maxLen + 2, 40) };
            });
        }
        XLSX.utils.book_append_sheet(wb, ws, data.sheetName || 'フォームの回答');

        const today = new Date();
        const fileName = `フォーム回答_${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}.xlsx`;
        XLSX.writeFile(wb, fileName);
    } catch (e) {
        console.error("フォーム回答DLエラー:", e);
        alert('ダウンロードに失敗しました: ' + (e.message || '不明なエラー'));
    } finally {
        hideLoading();
    }
}

// ====== Excelダウンロード ======
function downloadExcel() {
    if (!appData || Object.keys(appData).length === 0) {
        alert('データがありません。');
        return;
    }
    
    const wb = XLSX.utils.book_new();
    
    // 全日程を1シートにまとめる「全体一覧」シート
    const summaryRows = [];
    summaryRows.push(['日程', '氏名', '学年', '受験級', 'タブレット', '出欠', 'ReadPass (年/回)', 'Rスコア', 'AudiPass (年/回)', 'Lスコア', '備考・個別宿題']);
    
    sessionsInfo.forEach(session => {
        const sessionData = appData[session.id];
        if (!sessionData || !sessionData.participants) return;
        
        Object.keys(sessionData.participants).forEach(pid => {
            const pData = sessionData.participants[pid];
            const p = participantsList.find(x => x.id === pid);
            if (!p) return;
            
            summaryRows.push([
                session.date,
                p.name,
                p.schoolYear || '',
                p.grade || '',
                p.hasTablet ? '持参' : 'なし',
                pData.attended ? '出席' : '欠席',
                pData.rpContent || '',
                pData.rpScore || '',
                pData.apContent || '',
                pData.apScore || '',
                pData.remarks || ''
            ]);
        });
    });
    
    const summaryWs = XLSX.utils.aoa_to_sheet(summaryRows);
    // 列幅を設定
    summaryWs['!cols'] = [
        { wch: 14 }, // 日程
        { wch: 14 }, // 氏名
        { wch: 8 },  // 学年
        { wch: 10 }, // 受験級
        { wch: 8 },  // タブ
        { wch: 6 },  // 出欠
        { wch: 22 }, // ReadPass
        { wch: 8 },  // Rスコア
        { wch: 22 }, // AudiPass
        { wch: 8 },  // Lスコア
        { wch: 24 }, // 備考
    ];
    XLSX.utils.book_append_sheet(wb, summaryWs, '全体一覧');
    
    // 各日程ごとのシートを作成
    sessionsInfo.forEach(session => {
        const sessionData = appData[session.id];
        if (!sessionData || !sessionData.participants) return;
        
        const rows = [];
        rows.push(['氏名', '学年', '受験級', 'タブレット', '出欠', 'ReadPass (年/回)', 'Rスコア', 'AudiPass (年/回)', 'Lスコア', '備考・個別宿題']);
        
        Object.keys(sessionData.participants).forEach(pid => {
            const pData = sessionData.participants[pid];
            const p = participantsList.find(x => x.id === pid);
            if (!p) return;
            
            rows.push([
                p.name,
                p.schoolYear || '',
                p.grade || '',
                p.hasTablet ? '持参' : 'なし',
                pData.attended ? '出席' : '欠席',
                pData.rpContent || '',
                pData.rpScore || '',
                pData.apContent || '',
                pData.apScore || '',
                pData.remarks || ''
            ]);
        });
        
        // 全体宿題・メモ
        rows.push([]);
        rows.push(['全体宿題', sessionData.generalHomework || '']);
        rows.push(['日誌・引継ぎ', sessionData.generalNotes || '']);
        
        const ws = XLSX.utils.aoa_to_sheet(rows);
        ws['!cols'] = [
            { wch: 14 }, { wch: 8 }, { wch: 10 }, { wch: 8 }, { wch: 6 },
            { wch: 22 }, { wch: 8 }, { wch: 22 }, { wch: 8 }, { wch: 24 }
        ];
        // シート名は31文字制限 & 不正文字除去
        const sheetName = session.date.replace(/[\[\]\*\?\/\\]/g, '').substring(0, 31);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    
    // ダウンロード
    const today = new Date();
    const fileName = `英検勉強会_日誌_${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

// ====== 診断機能（同期確認用） ======
async function diagnoseCloudSync() {
    const _normName = (s) => String(s || '').replace(/[\s　]+/g, ' ').trim();
    const lines = [];
    lines.push('=== クラウド同期 診断レポート ===');
    lines.push('時刻: ' + new Date().toLocaleString('ja-JP'));
    lines.push('');

    // 1. GET で現在のクラウド状態を取得
    lines.push('【1】クラウドからのGET取得');
    let getData = null;
    try {
        const url = GAS_WEB_APP_URL + (GAS_WEB_APP_URL.includes('?') ? '&' : '?') + 't=' + Date.now();
        const res = await fetch(url, { cache: 'no-store' });
        getData = await res.json();
        lines.push('  ✓ GET成功 (HTTP ' + res.status + ')');
        lines.push('  クラウドの最終更新: ' + (getData.lastUpdated || '(未保存)'));
        if (getData._meta) {
            lines.push('  保存用シート存在: ' + getData._meta.storageSheetExists);
            lines.push('  参加者(クラウド): ' + getData._meta.participantsListCount + '名');
            lines.push('  日程(クラウド): ' + getData._meta.sessionsInfoCount + '件');
            lines.push('  日程データ(appData): ' + getData._meta.appDataSessionCount + 'セッション');
        } else {
            lines.push('  ⚠ _metaが無い → GAS旧バージョン(再デプロイ必須!)');
        }
        lines.push('  フォーム回答件数: ' + (getData.formResponses || []).length);
    } catch (e) {
        lines.push('  ✗ GET失敗: ' + e.message);
    }
    lines.push('');

    // 2. 現在のローカル状態
    lines.push('【2】このデバイスのローカル状態');
    lines.push('  参加者(ローカル): ' + participantsList.length + '名');
    lines.push('  日程(ローカル): ' + sessionsInfo.length + '件');
    lines.push('  日程データ: ' + Object.keys(appData).length + 'セッション');
    lines.push('  削除済みリスト: ' + (deletedParticipantNames || []).length + '名');
    lines.push('');

    // 3. ★ フォームと参加者リスト・削除済みリストの整合性チェック
    lines.push('【3】フォーム⇔アプリの整合性チェック');
    if (getData && Array.isArray(getData.formResponses)) {
        const formNamesSet = new Set();
        getData.formResponses.forEach(row => {
            const nameKey = Object.keys(row).find(k => k.includes('氏名'));
            if (nameKey) {
                const n = _normName(row[nameKey]);
                if (n) formNamesSet.add(n);
            }
        });
        const formNames = Array.from(formNamesSet);
        const listedNames = new Set(participantsList.map(p => _normName(p.name)));
        const deletedSet = new Set((deletedParticipantNames || []).map(_normName));

        lines.push('  フォームのユニーク氏名: ' + formNames.length + '名');
        lines.push('  participantsListに存在: ' + participantsList.length + '名');
        lines.push('');

        // フォームにいるのに participantsList にいない人
        const missingFromList = formNames.filter(n => !listedNames.has(n));
        // その中で削除済みリストに入っている人（＝ブロックされている）
        const blockedByDeleted = missingFromList.filter(n => deletedSet.has(n));
        // 削除済みリストに入っているが、フォームには存在する人（復活すべき候補）
        const shouldRecover = formNames.filter(n => deletedSet.has(n));

        if (missingFromList.length > 0) {
            lines.push('  ⚠ フォームにいるがアプリに反映されていない人:');
            missingFromList.forEach(n => {
                const flag = deletedSet.has(n) ? ' ← 【削除済みリストでブロック中】' : '';
                lines.push('    ・' + n + flag);
            });
        } else {
            lines.push('  ✓ フォーム氏名はすべてアプリに反映されています');
        }
        lines.push('');

        if (shouldRecover.length > 0) {
            lines.push('  🔧 復旧可能（削除済みリストから外せば再登場します）:');
            shouldRecover.forEach(n => lines.push('    ・' + n));
            lines.push('');
            lines.push('  → サイドバー下の「削除リストをリセット」ボタンで一括復旧できます');
        }

        if ((deletedParticipantNames || []).length > 0) {
            lines.push('');
            lines.push('  削除済みリストの全内容:');
            (deletedParticipantNames || []).forEach(n => lines.push('    ・' + n));
        }
    } else {
        lines.push('  フォーム回答がクラウドから取得できなかったためスキップ');
    }
    lines.push('');

    const report = lines.join('\n');
    console.log(report);
    alert(report);
    return report;
}

// ====== 削除済みリストから選択的に復帰 ======
async function openRestoreModal() {
    const list = deletedParticipantNames || [];
    if (list.length === 0) {
        alert('削除済みリストは既に空です。');
        return;
    }

    // フォーム回答の最新氏名を取得（「フォームにまだ存在する＝復帰で日誌に再登場する」の判定用）
    showLoading('フォーム回答を取得中...');
    let formNames = new Set();
    try {
        const url = GAS_WEB_APP_URL + (GAS_WEB_APP_URL.includes('?') ? '&' : '?') + 't=' + Date.now();
        const res = await fetch(url, { cache: 'no-store' });
        const data = await res.json();
        const _normName = (s) => String(s || '').replace(/[\s　]+/g, ' ').trim();
        if (Array.isArray(data.formResponses)) {
            data.formResponses.forEach(row => {
                const nameKey = Object.keys(row).find(k => k.includes('氏名'));
                if (nameKey) {
                    const n = _normName(row[nameKey]);
                    if (n) formNames.add(n);
                }
            });
        }
    } catch (e) {
        console.warn('フォーム回答の取得に失敗:', e);
    } finally {
        hideLoading();
    }

    // モーダル内にチェックボックス付きのリストを描画
    const container = document.getElementById('restoreList');
    container.innerHTML = '';
    list.forEach((name, idx) => {
        const inForm = formNames.has(name);
        const row = document.createElement('label');
        row.className = 'restore-item' + (inForm ? ' in-form' : '');
        row.innerHTML = `
            <input type="checkbox" class="restore-check" data-index="${idx}" ${inForm ? 'checked' : ''}>
            <span class="restore-name">${name}</span>
            <span class="restore-badge">${inForm ? '📥 フォームに存在（復帰で日誌に再登場）' : '📭 フォームからも削除済み（復帰しても日誌に戻らない）'}</span>
        `;
        container.appendChild(row);
    });

    document.getElementById('restoreCount').textContent = `計 ${list.length} 名`;
    document.getElementById('restoreModal').style.display = 'flex';
}

function closeRestoreModal() {
    document.getElementById('restoreModal').style.display = 'none';
}

function toggleAllRestoreSelection(checked) {
    document.querySelectorAll('#restoreList .restore-check').forEach(cb => {
        cb.checked = checked;
    });
}

async function executeRestore() {
    const checks = document.querySelectorAll('#restoreList .restore-check');
    const selectedIndices = [];
    checks.forEach(cb => {
        if (cb.checked) selectedIndices.push(parseInt(cb.dataset.index));
    });
    if (selectedIndices.length === 0) {
        alert('復帰する氏名を選択してください。');
        return;
    }

    const selectedSet = new Set(selectedIndices);
    const restoredNames = [];
    const remaining = [];
    (deletedParticipantNames || []).forEach((name, idx) => {
        if (selectedSet.has(idx)) {
            restoredNames.push(name);
        } else {
            remaining.push(name);
        }
    });

    deletedParticipantNames = remaining;
    localStorage.setItem('eikenDaigakumaeDeleted', JSON.stringify(deletedParticipantNames));

    closeRestoreModal();

    try {
        await pushStateToCloud();
        await syncWithCloud();
        alert(
            `${restoredNames.length} 名を削除済みリストから復帰しました:\n` +
            restoredNames.map(n => '・' + n).join('\n') +
            `\n\nフォームに存在する人は日誌に再登場しています。日誌を確認してください。`
        );
    } catch (e) {
        alert('復帰しましたが、クラウド同期でエラーが発生しました:\n' + (e.message || e));
    }
}

// 起動
init();
