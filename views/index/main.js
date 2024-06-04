function toggleExecute() {
    var executeBtn = $("#execute-btn");
    var loginSuccessBtn = $("#login-success-btn"); // 定位到登录账号成功按钮
    var startSuccessBtn = $("#start-success-btn"); // 定位到登录账号成功按钮

    if (executeBtn.text() === "开始执行") {
        start(executeBtn, loginSuccessBtn, startSuccessBtn)

    } else {
        // 检查按钮是否为灰色（即检查是否有btn-secondary类）
        if (loginSuccessBtn.hasClass("btn-secondary")) {
            pywebview.api.stop().then(function (response) {
                executeBtn.text("开始执行");
                executeBtn.removeClass("btn-danger").addClass("btn-primary");
                loginSuccessBtn.removeClass("btn-success").addClass("btn-secondary"); // 再变回灰色
                startSuccessBtn.removeClass("btn-success").addClass("btn-secondary"); // 再变回灰色
            })
            return
        }

        alert("请直接关闭浏览器！后点击登录账号成功！")

    }
}

// 初始化富文本编辑器
const editorConfig = {
    placeholder: '在此处输入...',
    onChange(editor) {
        console.log('内容变更:', editor.getHtml())
    }
};

const editor = window.wangEditor.createEditor({
    selector: '#editor-container',
    config: editorConfig
});

const toolbar = window.wangEditor.createToolbar({
    editor,
    selector: '#toolbar-container > div',
    config: {}
});

function exportHTML() {
    const htmlContent = editor.getHtml();
    const blob = new Blob([htmlContent], {type: "text/html"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'editor-content.html';
    document.body.appendChild(a);
    a.click();
    a.remove();
}

function start(executeBtn, loginSuccessBtn, startSuccessBtn) {
    // var searchKeyword = $('input[placeholder="请输入搜索关键字"]').val();
    var startPage = $('input[placeholder="从第几页开始"]').val();

    // if (searchKeyword.trim().length < 2) {
    //     alert('搜索关键字必须大于2个字符');
    //     return
    // }

    if (startPage.trim() === '') {
        alert('请输入搜索关键字和开始页');
        return
    }

    searchKeyword = " "
    pywebview.api.start(searchKeyword, startPage).then(function (response) {
        if (response === "running") {
            alert("正在执行中");
        }
    })

    executeBtn.text("停止执行");
    executeBtn.removeClass("btn-primary").addClass("btn-danger");
    loginSuccessBtn.removeClass("btn-secondary").addClass("btn-success"); // 变为绿色
    startSuccessBtn.removeClass("btn-secondary").addClass("btn-success"); // 变为绿色
}

function loginSuccess() {
    var loginBtn = $("#login-success-btn");

    // 检查按钮是否为灰色（即检查是否有btn-secondary类）
    if (loginBtn.hasClass("btn-secondary")) {
        console.log("按钮是灰色的，方法不执行");
        return; // 如果是灰色，这个方法不执行任何操作
    }

    pywebview.api.login_success().then(function (response) {
        var loginSuccessBtn = $("#login-success-btn"); // 定位到登录账号成功按钮
        loginSuccessBtn.removeClass("btn-success").addClass("btn-secondary"); // 变为灰色
    })
}

function startSuccess() {
    var loginBtn = $("#start-success-btn");

    // 检查按钮是否为灰色（即检查是否有btn-secondary类）
    if (loginBtn.hasClass("btn-secondary")) {
        console.log("按钮是灰色的，方法不执行");
        return; // 如果是灰色，这个方法不执行任何操作
    }

    pywebview.api.start_success().then(function (response) {
        var loginSuccessBtn = $("#start-success-btn"); // 定位到登录账号成功按钮
        loginSuccessBtn.removeClass("btn-success").addClass("btn-secondary"); // 变为灰色
    })
}

function stop() {
    var executeBtn = $("#execute-btn");
    var loginSuccessBtn = $("#login-success-btn"); // 定位到登录账号成功按钮
    var startSuccessBtn = $("#start-success-btn"); // 定位到登录账号成功按钮
    executeBtn.text("开始执行");
    executeBtn.removeClass("btn-danger").addClass("btn-primary");
    loginSuccessBtn.removeClass("btn-success").addClass("btn-secondary"); // 再变回灰色
    startSuccessBtn.removeClass("btn-success").addClass("btn-secondary"); // 再变回灰色
}


function send_email() {

    // 使用jQuery获取输入框的值
    var senderEmail = $('input[placeholder="发件人邮箱"]').val();
    var emailAuthCode = $('input[placeholder="邮箱授权码"]').val();
    var emailSubject = $('input[placeholder="邮件主题"]').val();

    // 使用jQuery获取id="lab"中所有<p>元素的文本内容，并存储在一个数组中
    var labTexts = $('#lab p span').map(function () {
        return $(this).text();
    }).get(); // 使用.get()将jQuery对象转换为数组

    // 判空操作
    if (!senderEmail.trim()) {
        alert('发件人邮箱不能为空');
        return;
    }
    if (!emailAuthCode.trim()) {
        alert('邮箱授权码不能为空');
        return;
    }
    if (!emailSubject.trim()) {
        alert('邮件主题不能为空');
        return;
    }

    // 如果所有输入都已填写，以下代码将处理发送邮件的逻辑
    html_text = get_html()

    // 在这里添加实际的邮件发送逻辑

    console.log(senderEmail)
    console.log(emailAuthCode)
    console.log(emailSubject)
    console.log(html_text)
    console.log(labTexts)
    pywebview.api.start_send(senderEmail, emailAuthCode, emailSubject, html_text, labTexts).then(function (response) {
        alert("发送成功");
    })
}

function set_file() {
    pywebview.api.set_file().then(function (response) {
        // 创建包含响应文本的p元素
        var $textSpan = $('<span>').text(response); // 创建一个span来存放文本内容
        var $p = $('<p>');
        // 创建删除按钮，并应用Bootstrap的样式
        var $deleteBtn = $('<button>')
            .text('删除')
            .addClass('btn btn-danger btn-sm') // 使用Bootstrap类来设置按钮的样式
            .css('margin-left', '10px');

        // 为删除按钮添加点击事件处理程序，删除p元素
        $deleteBtn.on('click', function () {
            $p.remove();
        });

        // 将文本span和删除按钮附加到p元素
        $p.append($textSpan).append($deleteBtn);
        // 将p元素附加到id="lab"的div中
        $('#lab').append($p);
    });
}

function get_html() {
    // 获取编辑器的HTML内容
    return editor.getHtml(); // 直接返回HTML内容
}
