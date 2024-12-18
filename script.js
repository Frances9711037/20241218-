// 設定 Excel 檔案的 URL（請替換成您的 Excel 檔案路徑）
const EXCEL_FILE_URL = 'quiz_data.xlsx';

let quizData = {
    questions: []
};
let currentQuestionIndex = 0;
let userAnswers = [];

// 載入 Excel 檔案
async function loadExcelFile() {
    try {
        const response = await fetch(EXCEL_FILE_URL);
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // 假設第一個工作表包含題目
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // 轉換 Excel 數據為題目格式
        quizData.questions = jsonData.map(row => {
            const question = {
                type: row.type,
                question: row.question,
                correctAnswer: row.correctAnswer.toString().trim() // 確保正確答案是字串
            };
            
            if (row.type === 'single' || row.type === 'multiple') {
                if (typeof row.options === 'string') {
                    question.options = row.options.split(',').map(opt => opt.trim());
                } else {
                    question.options = row.options.toString().split(',').map(opt => opt.trim());
                }

                if (row.type === 'multiple') {
                    question.correctAnswer = question.correctAnswer.split(',').map(ans => ans.trim());
                }
            }
            
            return question;
        });

        userAnswers = new Array(quizData.questions.length).fill(null);
        generateCurrentQuestion();
        startTimer(60);
        updateProgress();
        
    } catch (error) {
        console.error('載入題目失敗:', error);
        alert('載入題目失敗，請稍後再試');
    }
}

// 生成當前題目
function generateCurrentQuestion() {
    const quizContainer = document.getElementById('quiz-container');
    quizContainer.innerHTML = '';
    
    if (currentQuestionIndex >= quizData.questions.length) {
        return;
    }
    
    const q = quizData.questions[currentQuestionIndex];
    const questionDiv = document.createElement('div');
    questionDiv.className = 'question-container';
    
    const questionText = document.createElement('p');
    questionText.textContent = `${currentQuestionIndex + 1}. ${q.question}`;
    questionDiv.appendChild(questionText);
    
    if (q.type === 'single' || q.type === 'multiple') {
        const optionsDiv = document.createElement('div');
        optionsDiv.className = 'options-container';
        
        q.options.forEach(option => {
            const label = document.createElement('label');
            const input = document.createElement('input');
            input.type = q.type === 'single' ? 'radio' : 'checkbox';
            input.name = `question${currentQuestionIndex}`;
            input.value = option;
            
            if (userAnswers[currentQuestionIndex]) {
                if (q.type === 'single') {
                    if (userAnswers[currentQuestionIndex] === option) {
                        input.checked = true;
                    }
                } else {
                    if (userAnswers[currentQuestionIndex].includes(option)) {
                        input.checked = true;
                    }
                }
            }
            
            label.appendChild(input);
            label.appendChild(document.createTextNode(option));
            optionsDiv.appendChild(label);
        });
        
        questionDiv.appendChild(optionsDiv);
    } else if (q.type === 'fill') {
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'fill-in';
        input.name = `question${currentQuestionIndex}`;
        input.placeholder = '請在此輸入答案';
        
        if (userAnswers[currentQuestionIndex]) {
            input.value = userAnswers[currentQuestionIndex];
        }
        
        questionDiv.appendChild(input);
    }
    
    quizContainer.appendChild(questionDiv);
    updateProgress();
}

// 保存當前答案
function saveCurrentAnswer() {
    const q = quizData.questions[currentQuestionIndex];
    
    if (q.type === 'single') {
        const selected = document.querySelector(`input[name="question${currentQuestionIndex}"]:checked`);
        userAnswers[currentQuestionIndex] = selected ? selected.value : null;
    } else if (q.type === 'multiple') {
        const selected = Array.from(document.querySelectorAll(`input[name="question${currentQuestionIndex}"]:checked`))
            .map(input => input.value);
        userAnswers[currentQuestionIndex] = selected;
    } else if (q.type === 'fill') {
        const answer = document.querySelector(`input[name="question${currentQuestionIndex}"]`).value.trim();
        userAnswers[currentQuestionIndex] = answer;
    }
}

// 計時器功能
function startTimer(minutes) {
    const timerDisplay = document.querySelector('#timer span');
    let totalSeconds = minutes * 60;
    
    const timer = setInterval(() => {
        const minutesLeft = Math.floor(totalSeconds / 60);
        const secondsLeft = totalSeconds % 60;
        
        timerDisplay.textContent = `${String(minutesLeft).padStart(2, '0')}:${String(secondsLeft).padStart(2, '0')}`;
        
        if (totalSeconds <= 0) {
            clearInterval(timer);
            submitQuiz();
        }
        
        totalSeconds--;
    }, 1000);
}

// 更新進度
function updateProgress() {
    const progressDisplay = document.querySelector('#progress span');
    progressDisplay.textContent = `${currentQuestionIndex + 1}/${quizData.questions.length}`;
}

// 計算分數
function calculateScore() {
    let score = 0;
    const totalQuestions = quizData.questions.length;
    
    quizData.questions.forEach((q, index) => {
        const userAns = userAnswers[index];
        if (!userAns) return; // 如果沒有答案，跳過

        if (q.type === 'single') {
            // 確保比較的是字串
            if (userAns.toString().trim() === q.correctAnswer.toString().trim()) {
                score++;
            }
        } else if (q.type === 'multiple') {
            const correctAns = Array.isArray(q.correctAnswer) ? 
                q.correctAnswer : 
                q.correctAnswer.split(',').map(ans => ans.trim());
            
            if (JSON.stringify(userAns.sort()) === JSON.stringify(correctAns.sort())) {
                score++;
            }
        } else if (q.type === 'fill') {
            const correctAnswers = q.correctAnswer.split('|').map(ans => ans.trim().toLowerCase());
            if (correctAnswers.includes(userAns.toLowerCase())) {
                score++;
            }
        }
    });
    
    return (score / totalQuestions) * 100;
}

// 檢查答案
function checkAnswer(question, userAnswer, index) {
    if (!userAnswer) return false;

    if (question.type === 'single') {
        return userAnswer.toString().trim() === question.correctAnswer.toString().trim();
    } else if (question.type === 'multiple') {
        const correctAnswers = Array.isArray(question.correctAnswer) ? 
            question.correctAnswer.sort() : 
            question.correctAnswer.split(',').map(ans => ans.trim()).sort();
        const userAnswers = (userAnswer || []).sort();
        return JSON.stringify(userAnswers) === JSON.stringify(correctAnswers);
    } else if (question.type === 'fill') {
        const correctAnswers = question.correctAnswer.split('|').map(ans => ans.trim().toLowerCase());
        return correctAnswers.includes(userAnswer.toLowerCase());
    }
    return false;
}

// 提交測驗
function submitQuiz() {
    saveCurrentAnswer();
    const score = calculateScore();
    
    document.getElementById('quiz-container').style.display = 'none';
    document.getElementById('next-btn').style.display = 'none';
    document.getElementById('submit-btn').style.display = 'none';

    showResult(score);
    showAnswerReview();
}

// 顯示結果
function showResult(score) {
    const resultDiv = document.getElementById('result');
    resultDiv.style.display = 'block';
    resultDiv.innerHTML = `<h2>測驗結果</h2>
                          <p class="score">得分：${score.toFixed(1)}分</p>`;
    
    if (score < 60) {
        resultDiv.innerHTML += '<p class="message low-score">再加油！</p>';
    } else if (score >= 90) {
        resultDiv.innerHTML += '<p class="message high-score">太棒了！</p>';
    }
}

// 顯示答案對照
function showAnswerReview() {
    const reviewContainer = document.getElementById('answer-review');
    reviewContainer.style.display = 'block';
    reviewContainer.innerHTML = '<h2>答案對照</h2>';
    
    quizData.questions.forEach((q, index) => {
        const reviewItem = document.createElement('div');
        reviewItem.className = 'review-item';
        
        reviewItem.innerHTML = `<div class="question-text"><strong>題目 ${index + 1}:</strong> ${q.question}</div>`;
        
        let userAnswerText = '��作答';
        if (userAnswers[index]) {
            if (q.type === 'multiple') {
                userAnswerText = Array.isArray(userAnswers[index]) ? 
                    userAnswers[index].join(', ') : userAnswers[index];
            } else {
                userAnswerText = userAnswers[index];
            }
        }
        
        let correctAnswerText = '';
        if (q.type === 'multiple') {
            correctAnswerText = Array.isArray(q.correctAnswer) ? 
                q.correctAnswer.join(', ') : 
                q.correctAnswer.split(',').map(ans => ans.trim()).join(', ');
        } else {
            correctAnswerText = q.correctAnswer;
        }
        
        const isCorrect = checkAnswer(q, userAnswers[index], index);
        
        reviewItem.innerHTML += `
            <div class="user-answer">你的答案: <span class="${isCorrect ? 'correct' : 'incorrect'}">${userAnswerText}</span></div>
            <div class="correct-answer">正確答案: ${correctAnswerText}</div>
        `;
        
        reviewContainer.appendChild(reviewItem);
    });
}

// 初始化
window.onload = function() {
    loadExcelFile();
    
    document.getElementById('next-btn').addEventListener('click', () => {
        saveCurrentAnswer();
        currentQuestionIndex++;
        
        if (currentQuestionIndex >= quizData.questions.length) {
            document.getElementById('next-btn').style.display = 'none';
            document.getElementById('submit-btn').style.display = 'block';
            currentQuestionIndex = quizData.questions.length - 1;
        }
        
        generateCurrentQuestion();
    });
    
    document.getElementById('submit-btn').addEventListener('click', submitQuiz);
}; 