const SHEET_API_URL = 
'https://script.google.com/macros/s/AKfycbwdxIaL5ypL1rjiub-0S2Xak1tO0sdlOzGxZcA9WyKJVh468kRN5QmFqU-a6ojTDwMP/exec';



function fetchOnlineSheetData() {
    return new Promise((resolve, reject) => {
        const callbackName = 'handleSheetCallback_' + Date.now();
        const script = document.createElement('script');
        const timeout = setTimeout(() => {
            cleanup();
            reject(new Error('응답 시간 초과'));
        }, 10000);

        function cleanup() {
            clearTimeout(timeout);
            delete window[callbackName];
            if (script.parentNode) script.parentNode.removeChild(script);
        }

        window[callbackName] = function(json) {
            cleanup();
            if (!json) {
                reject(new Error('빈 응답'));
            } else if (json.error) {
                reject(new Error(json.error));
            } else {
                resolve(json.items || []);
            }
        };

        script.onerror = function() {
            cleanup();
            reject(new Error('스크립트 로드 실패 (배포 권한 및 URL 확인 필요)'));
        };

        const sep = SHEET_API_URL.includes('?') ? '&' : '?';
        script.src = `${SHEET_API_URL}${sep}callback=${callbackName}&_t=${Date.now()}`;
        document.body.appendChild(script);
    });
}

function formatKRW(num) {
    return Number(num || 0).toLocaleString('ko-KR') + ' 원';
}

function calcIndexReturns(items, fieldKey, startIndex = 0) {
    if (!items || items.length === 0 || startIndex < 0 || startIndex >= items.length) {
        return [];
    }
    const baseVal = items[startIndex][fieldKey];
    if (baseVal === null || baseVal === undefined || Number(baseVal) === 0) {
        return items.map(() => null);
    }
    const baseNum = Number(baseVal);
    return items.map(item => {
        const curr = item[fieldKey];
        if (curr === null || curr === undefined || Number(curr) === 0) return null;
        return Number((((Number(curr) - baseNum) / baseNum) * 100).toFixed(2));
    });
}