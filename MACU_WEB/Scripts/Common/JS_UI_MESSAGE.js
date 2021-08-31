// 新增修改成功時，跳出訊息
function OnSuccess(result) {
    //alert(strCookieCultureValue);

    var strRetrunPage = "回查詢畫面";
    var strContinue = "繼續調整";

    switch (strCookieCultureValue) {
        case "zh-TW":
            strRetrunPage = "回查詢畫面";
            strContinue = "繼續調整";
            break;
        case "zh-CN":
            strRetrunPage = "回查询画面";
            strContinue = "继续";
            break;
        case "en-US":
            strRetrunPage = "Back To Search Page";
            strContinue = "Continue";
            break;
        case "ja-JP":
            strRetrunPage = "調べ画面戻る";
            strContinue = "つづく";
            break;
        default:
            strRetrunPage = "回查詢畫面";
            strContinue = "繼續調整";
            break;
    }

    BootstrapDialog.show({
        title: result.title,
        message: result.Message + "!!",
        buttons: [{
            label: strRetrunPage,
            action: function (dialogRef) {
                window.location = result.returnAction;
                dialogRef.close();
            }
        }, {
            label: strContinue,
            action: function (dialogRef) {
                window.location = result.continueAction;
                dialogRef.close();
            }
        }]
    });
}
function OnSuccess_View(result) {
    //alert(strCookieCultureValue);

    var strRetrunPage = "回查詢畫面";
    var strContinue = "繼續調整";
    var strView = "查看單據";

    switch (strCookieCultureValue) {
        case "zh-TW":
            strRetrunPage = "回查詢畫面";
            strContinue = "繼續調整";
            strView = "查看單據";
            break;
        case "zh-CN":
            strRetrunPage = "回查询画面";
            strContinue = "继续";
            strView = "查看单据";
            break;
        case "en-US":
            strRetrunPage = "Back To Search Page";
            strContinue = "Continue";
            strView = "View";
            break;
        case "ja-JP":
            strRetrunPage = "調べ画面戻る";
            strContinue = "つづく";
            strView = "";
            break;
        default:
            strRetrunPage = "回查詢畫面";
            strContinue = "繼續調整";
            strView = "查看單據";
            break;
    }

    BootstrapDialog.show({
        title: result.title,
        message: result.Message + "!!",
        buttons: [{
            label: strRetrunPage,
            action: function (dialogRef) {
                window.location = result.returnAction;
                dialogRef.close();
            }
        }, {
            label: strView,
            action: function (dialogRef) {
                window.location = result.viewAction;
                dialogRef.close();
            }
        }, {
            label: strContinue,
            action: function (dialogRef) {
                window.location = result.continueAction;
                dialogRef.close();
            }
        }]
    });
}
// 新增修改成功時，跳出訊息，返回查詢畫面
function OnSuccess_Back(result) {
    //alert(strCookieCultureValue);

    var strRetrunPage = "回查詢畫面";
    
    switch (strCookieCultureValue) {
        case "zh-TW":
            strRetrunPage = "回查詢畫面";            
            break;
        case "zh-CN":
            strRetrunPage = "回查询画面";
            break;
        case "en-US":
            strRetrunPage = "Back To Search Page";
            break;
        case "ja-JP":
            strRetrunPage = "調べ画面戻る";
            break;
        default:
            strRetrunPage = "回查詢畫面";
            break;
    }

    BootstrapDialog.show({
        title: result.title,
        message: result.Message + "!!",
        buttons: [{
            label: strRetrunPage,
            action: function (dialogRef) {
                window.location = result.returnAction;
                dialogRef.close();
            }
        }]
    });
}


// 新增修改失敗時，跳出訊息
function OnFailure(result) {

    var strConfirm = "確定";
    top.console.log(result.title);
    top.console.log(result.Message);

    switch (strCookieCultureValue) {
        case "zh-TW":
            strConfirm = "確定";
            break;
        case "zh-CN":
            strConfirm = "确定";
            break;
        case "en-US":
            strConfirm = "Confirm";
            break;
        case "ja-JP":
            strConfirm = "決定します";
            break;
        default:
            strConfirm = "確定";
            break;
    }

    BootstrapDialog.alert({
        title: result.title,
        message: result.Message + "!!",
        type: BootstrapDialog.TYPE_DANGER,
        buttonLabel: strConfirm
    });
}

// bootstrap alert 訊息
function JsFunAlert(e) {

    var strConfirm = "確定";

    switch (strCookieCultureValue) {
        case "zh-TW":
            strConfirm = "確定";
            break;
        case "zh-CN":
            strConfirm = "确定";
            break;
        case "en-US":
            strConfirm = "Confirm";
            break;
        case "ja-JP":
            strConfirm = "決定します";
            break;
        default:
            strConfirm = "確定";
            break;
    }

    BootstrapDialog.show({
        title: e.title,
        message: e.message,
        buttons: [{
            label: strConfirm,
            action: function (dialogRef) {
                dialogRef.close();
            }
        }]
    });
}

//20210129 CCL+ ////簡單版沒有button action /////
// bootstrap alert 訊息
function JsFunAlertMsg(e) {

    var strConfirm = "確定";

    switch (strCookieCultureValue) {
        case "zh-TW":
            strConfirm = "確定";
            break;
        case "zh-CN":
            strConfirm = "确定";
            break;
        case "en-US":
            strConfirm = "Confirm";
            break;
        case "ja-JP":
            strConfirm = "決定します";
            break;
        default:
            strConfirm = "確定";
            break;
    }

    BootstrapDialog.alert({
        title: e.title,
        message: e.message,
        //type: BootstrapDialog.TYPE_DANGER,
        buttonLabel: strConfirm
    });
}
//////////////////////////////////////////////

// 新增修改成功時，跳出訊息
function JsFunConfirm(e) {

    var strConfirm = "確定";
    var strCancel = "取消";

    switch (strCookieCultureValue) {
        case "zh-TW":
            strConfirm = "確定";
            strCancel = "取消";
            break;
        case "zh-CN":
            strConfirm = "确定";
            strCancel = "取消";
            break;
        case "en-US":
            strConfirm = "Confirm";
            strCancel = "Cancel";
            break;
        case "ja-JP":
            strConfirm = "決定します";
            strCancel = "キャンセル";
            break;
        default:
            strConfirm = "確定";
            strCancel = "取消";
            break;
    }
    BootstrapDialog.show({
        title: e.title,
        message: e.Message + "!!",
        buttons: [{
            label: strConfirm,
            action: function (dialogRef) {
                dialogRef.close();
                return true;
            }
        }, {
            label: strCancel,
            action: function (dialogRef) {
                dialogRef.close();
                return false;
            }
        }]
    });
}

function JsFunErrorMsg(I_TITLE,I_MESSAGE,I_OBJ)
{
    var strConfirm = "確定";

    switch (strCookieCultureValue) {
        case "zh-TW":
            strConfirm = "確定";
            break;
        case "zh-CN":
            strConfirm = "确定";
            break;
        case "en-US":
            strConfirm = "Confirm";
            break;
        case "ja-JP":
            strConfirm = "決定します";
            break;
        default:
            strConfirm = "確定";
            break;
    }

    var types = [BootstrapDialog.TYPE_DANGER];

    BootstrapDialog.show({
        title: I_TITLE,
        message: I_MESSAGE,
        type: BootstrapDialog.TYPE_DANGER,
        buttons: [{
            label: strConfirm,           
            action: function (dialogRef) {
                dialogRef.close();
                if (I_OBJ !== null)
                $(I_OBJ).focus();
            }
        }]
    });
}