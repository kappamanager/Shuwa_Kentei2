document.addEventListener('DOMContentLoaded', function () {
    const levelSelect = document.getElementById('level');
    const themeButton = document.getElementById('theme-button');
    const themeElement = document.getElementById('theme');
    const timerSelect = document.getElementById('timer-select');
    const startButton = document.getElementById('start-button');
    const exampleButton = document.getElementById('example-button');
    const timerElement = document.getElementById('timer');
    let data = [];
    let timer;

    function loadExcelDataFromServer(filePath) {
        fetch(filePath)
            .then(response => response.arrayBuffer())
            .then(arrayBuffer => {
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                data = XLSX.utils.sheet_to_json(worksheet);
                console.log("データが読み込まれました。", data); // デバッグ用ログ
                populateLevels();
            })
            .catch(error => console.error("エラーが発生しました。", error)); // エラーハンドリング
    }

    function populateLevels() {
        const levels = [...new Set(data.map(item => item.Level))];
        console.log("レベルが一意に抽出されました。", levels); // デバッグ用ログ
        levelSelect.innerHTML = ""; // 既存のオプションをクリア
        levels.forEach(level => {
            const option = document.createElement('option');
            option.value = level;
            option.textContent = level;
            levelSelect.appendChild(option);
        });
    }

    function showTheme() {
        const selectedLevel = levelSelect.value;
        const filteredThemes = data.filter(item => item.Level == selectedLevel).map(item => item.Theme);
        const themeIndex = Math.floor(Math.random() * filteredThemes.length);
        themeElement.textContent = filteredThemes[themeIndex];
    }

    function startTimer() {
        let time = parseInt(timerSelect.value, 10) * 60;
        clearInterval(timer);
        timer = setInterval(() => {
            const minutes = Math.floor(time / 60);
            const seconds = time % 60;
            timerElement.textContent = `残り時間: ${minutes}:${seconds < 10 ? '0' : ''}${seconds}`;
            time--;
            if (time < 0) {
                clearInterval(timer);
                timerElement.textContent = '時間切れです。';
            }
        }, 1000);
        timerElement.classList.remove('hidden');
    }

    function showExample() {
        const exampleIndex = Math.floor(Math.random() * data.length);
        alert(data[exampleIndex].Theme);
    }

    themeButton.addEventListener('click', showTheme);
    startButton.addEventListener('click', startTimer);
    exampleButton.addEventListener('click', showExample);

    // サーバーからファイルをロード
    loadExcelDataFromServer('excel/shuwa_exam_question_list.xlsx');
});
