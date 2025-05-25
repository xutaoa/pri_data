// 全局变量存储分析结果
let analysisResults = {
    file1: null,
    file2: null,
    file3: null
};

// 初始化文件上传事件
document.addEventListener('DOMContentLoaded', function() {
    // 为三个文件输入框添加事件监听
    for (let i = 1; i <= 3; i++) {
        const fileInput = document.getElementById(`fileInput${i}`);
        fileInput.addEventListener('change', function(e) {
            handleFileUpload(e, i);
        });
    }
});

// 处理文件上传
function handleFileUpload(event, fileNumber) {
    const file = event.target.files[0];
    if (!file) return;
    
    const formData = new FormData();
    formData.append('file', file);
    
    const uploadBox = document.getElementById(`uploadBox${fileNumber}`);
    uploadBox.querySelector('button').textContent = '上传中...';
    uploadBox.querySelector('button').disabled = true;
    
    fetch('/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.status === 'success') {
            // 显示Sheet选择器
            const sheetSelector = document.getElementById(`sheetSelector${fileNumber}`);
            sheetSelector.classList.remove('d-none');
            
            // 填充Sheet选项
            const sheetSelect = document.getElementById(`sheetSelect${fileNumber}`);
            sheetSelect.innerHTML = '<option value="">选择Sheet</option>';
            data.sheets.forEach(sheet => {
                const option = document.createElement('option');
                option.value = sheet;
                option.textContent = sheet;
                sheetSelect.appendChild(option);
            });
            
            // 更新按钮状态
            uploadBox.querySelector('button').textContent = '上传成功';
            uploadBox.querySelector('button').classList.remove('btn-primary');
            uploadBox.querySelector('button').classList.add('btn-success');
            
            // 存储文件名
            analysisResults[`file${fileNumber}`] = {
                filename: data.filename,
                sheet: null,
                result: null
            };
        } else {
            alert(`上传失败: ${data.error}`);
            uploadBox.querySelector('button').textContent = `上传文件${fileNumber}`;
            uploadBox.querySelector('button').disabled = false;
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('上传过程中发生错误');
        uploadBox.querySelector('button').textContent = `上传文件${fileNumber}`;
        uploadBox.querySelector('button').disabled = false;
    });
}

// 分析数据
function analyze(fileNumber) {
    const sheetSelect = document.getElementById(`sheetSelect${fileNumber}`);
    const selectedSheet = sheetSelect.value;
    
    if (!selectedSheet) {
        alert('请选择一个Sheet');
        return;
    }
    
    // 更新全局变量
    analysisResults[`file${fileNumber}`].sheet = selectedSheet;
    
    // 发送分析请求
    fetch('/analyze', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            filename: analysisResults[`file${fileNumber}`].filename,
            sheet_name: selectedSheet,
            file_num: fileNumber
        })
    })
    .then(response => response.json())
    .then(data => {
        if (data.status === 'success') {
            // 存储分析结果
            analysisResults[`file${fileNumber}`].result = data.result;
            
            // 显示分析结果
            displayTextResult(fileNumber, data.result);
            
            // 更新预览区域
            updatePreview();
        } else {
            alert(`分析失败: ${data.error}`);
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('分析过程中发生错误');
    });
}

// 显示纯文本结果
function displayTextResult(fileNumber, result) {
    const analysisResultsDiv = document.getElementById('analysisResults');
    
    // 创建或更新结果卡片
    let resultCard = document.getElementById(`resultCard${fileNumber}`);
    if (!resultCard) {
        resultCard = document.createElement('div');
        resultCard.id = `resultCard${fileNumber}`;
        resultCard.className = 'card mb-3';
        analysisResultsDiv.appendChild(resultCard);
        
        const cardHeader = document.createElement('div');
        cardHeader.className = 'card-header d-flex justify-content-between align-items-center';
        cardHeader.innerHTML = `
            <h6>文件${fileNumber}分析结果</h6>
            <button class="btn btn-sm btn-outline-secondary" onclick="copyResult(${fileNumber})">
                <i class="bi bi-clipboard"></i> 复制
            </button>
        `;
        resultCard.appendChild(cardHeader);
        
        const cardBody = document.createElement('div');
        cardBody.className = 'card-body';
        resultCard.appendChild(cardBody);
    }
    
    // 填充结果内容
    const cardBody = resultCard.querySelector('.card-body');
    cardBody.innerHTML = `
        <p><strong>文件名:</strong> ${analysisResults[`file${fileNumber}`].filename}</p>
        <p><strong>Sheet名:</strong> ${analysisResults[`file${fileNumber}`].sheet}</p>
        <div class="text-result">
            <pre>${result}</pre>
        </div>
    `;
}

// 更新预览区域
function updatePreview() {
    const previewArea = document.getElementById('previewArea');
    
    // 检查是否所有文件都已分析
    let allAnalyzed = true;
    let combinedText = '';
    
    for (let i = 1; i <= 3; i++) {
        if (!analysisResults[`file${i}`] || !analysisResults[`file${i}`].result) {
            allAnalyzed = false;
            break;
        }
        if (i === 2){
            continue
        }
        combinedText += `${analysisResults[`file${i}`].result}\n\n`;
    }
    
    if (allAnalyzed) {
        previewArea.innerHTML = `<pre>${combinedText}</pre>`;
    } else {
        previewArea.innerHTML = '<p class="text-muted">请完成所有文件的分析</p>';
    }
}

// 复制单个结果
function copyResult(fileNumber) {
    const result = analysisResults[`file${fileNumber}`].result;
    if (result) {
        navigator.clipboard.writeText(result)
            .then(() => alert('已复制到剪贴板'))
            .catch(err => console.error('无法复制文本: ', err));
    }
}

// 复制预览区域内容
function copyPreview() {
    const previewText = document.getElementById('previewArea').textContent;
    if (previewText && !previewText.includes('分析结果将显示在这里')) {
        navigator.clipboard.writeText(previewText)
            .then(() => alert('已复制到剪贴板'))
            .catch(err => console.error('无法复制文本: ', err));
    }
}

// 复制所有结果
function copyAllResults() {
    let allText = '';
    for (let i = 1; i <= 3; i++) {
        if (analysisResults[`file${i}`] && analysisResults[`file${i}`].result) {
            allText += `\n${analysisResults[`file${i}`].result}\n\n`;
        }
    }
    
    if (allText) {
        navigator.clipboard.writeText(allText)
            .then(() => alert('已复制所有结果到剪贴板'))
            .catch(err => console.error('无法复制文本: ', err));
    }


    fetch('/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.status === 'success') {
            // 显示Sheet选择器
            const sheetSelector = document.getElementById(`sheetSelector${fileNumber}`);
            sheetSelector.classList.remove('d-none');
            
            // 填充Sheet选项
            const sheetSelect = document.getElementById(`sheetSelect${fileNumber}`);
            sheetSelect.innerHTML = '<option value="">选择Sheet</option>';
            data.sheets.forEach(sheet => {
                const option = document.createElement('option');
                option.value = sheet;
                option.textContent = sheet;
                sheetSelect.appendChild(option);
            });
            
            // 更新按钮状态
            uploadBox.querySelector('button').textContent = '上传成功';
            uploadBox.querySelector('button').classList.remove('btn-primary');
            uploadBox.querySelector('button').classList.add('btn-success');
            
            // 存储文件名
            analysisResults[`file${fileNumber}`] = {
                filename: data.filename,
                sheet: null,
                result: null
            };
        } else {
            alert(`上传失败: ${data.error}`);
            uploadBox.querySelector('button').textContent = `上传文件${fileNumber}`;
            uploadBox.querySelector('button').disabled = false;
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('上传过程中发生错误');
        uploadBox.querySelector('button').textContent = `上传文件${fileNumber}`;
        uploadBox.querySelector('button').disabled = false;
    });

}
