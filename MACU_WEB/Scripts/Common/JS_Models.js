/* ====================================================================================================
/ Bootstrap Model Dialog 提示跳窗 Style1: 垂直置中Model Dialog
/  開發人員:Tony Chen (20201209)
/ 範例: 顯示: JsFunModelS1_SHOW(true);
/       隱藏: JsFunModelS1_SHOW(false);
/ 參數: message: {
/                   title,
/                   body
/                }
/ ===================================================================================================== */
function JsFunModelS1_SHOW(e) {
    
    if (e == true) {
        //alert("AAA");
        $.modelDialogS1.show();        
    } else {
        $.modelDialogS1.hide();
    }
}
(function () { /* Bootstrap 載入中 Function */
    if (typeof ($) == 'function') {
        
        var waitModeling = waitModeling || (function ($) {
            'use strict';
            // Creating modal dialog's DOM
            var $dialog = $(
                '<div class="modal fade" data-backdrop="static" data-keyboard="false" tabindex="-1" role="dialog" aria-hidden="true" style="padding-top:15%; overflow-y:visible;">' +
                '<div class="modal-dialog modal-m">' +
                '<div class="modal-content">' +
                '<div class="modal-header"><h3 style="margin:0;"></h3></div>' +
                '<div class="modal-body">' +
                '<div class="progress progress-striped active" style="margin-bottom:0;"><div class="progress-bar" style="width: 100%"></div></div>' +
                '</div>' +
                '</div></div></div>'

                //< !--Modal -->
                '<div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">' +
                    '<div class="modal-dialog modal-dialog-centered" role="document">' +
                        '<div class="modal-content">' +
                            '<div class="modal-header">' +
                                '<h5 class="modal-title" id="exampleModalCenterTitle"></h5>' +
                                '<button type="button" class="close" data-dismiss="modal" aria-label="Close">' +
                                    '<span aria-hidden="true">&times;</span>' +
                                '</button>' +
                            '</div>' +
                            '<div class="modal-body">' +                            
                            '</div>' +
                            '<div class="modal-footer">' +
                                '<button type="button" class="btn btn-secondary" data-dismiss="modal">關閉</button>' +
                                '<button type="button" class="btn btn-primary">確認修改</button>' +
                            '</div>' +
                        '</div>' +
                    '</div>' +
                '</div>'
            );
            return {
                show: function (message, options) {
                    
                    // Assigning defaults
                    if (typeof options === 'undefined') {
                        options = {};
                    }
                    if (typeof (message) == 'object') {
                        options = message;
                        message = undefined;
                    }
                    if (typeof message === 'undefined') {
                        message = '載入中，請稍後...';
                        
                        /*
                        switch (strCookieCultureValue) {
                            case "zh-TW":
                                message = "載入中，請稍後...";
                                break;
                            case "zh-CN":
                                message = "载入中，请稍后...";
                                break;
                            case "en-US":
                                message = "Now Loading, Please wait...";
                                break;
                            case "ja-JP":
                                message = "ローディング，しばらくお待ちください...";
                                break;
                            default:
                                message = "載入中，請稍後...";
                                break;
                        }
                        */
                    }
                    //options = {}; // 不提供所有外擴, 一律預設值
                    var settings = $.extend({
                        dialogSize: 'm',
                        progressType: '',
                        onHide: null // This callback runs after the dialog was hidden
                    }, options);
                    
                    // Configuring dialog
                    //$dialog.find('.modal-dialog').attr('class', 'modal-dialog').addClass('modal-' + settings.dialogSize);
                    //$dialog.find('.progress-bar').attr('class', 'progress-bar');
                    //if (settings.progressType) {
                    //    $dialog.find('.progress-bar').addClass('progress-bar-' + settings.progressType);
                    //}
                    if (typeof (settings) === 'object') {
                        //message json title
                        $dialog.find('h5.modal-title').text(settings.title);
                        //message json body
                        $dialog.find('div.modal-body').text(settings.body);
                        //message json OK, Close Btn
                        //$dialog.find('div.modal-body').text(message.ok);
                    }

                    $dialog.find('h3').text(message);
                    // Adding callbacks
                    if (typeof settings.onHide === 'function') {
                        $dialog.off('hidden.bs.modal').on('hidden.bs.modal', function (e) {
                            settings.onHide.call($dialog);
                        });
                    }
                    /* Opening dialog */
                    
                    $dialog.modal();
                },
                /* Closes dialog */
                hide: function (options) {
                    var hideMS = 500;

                    if (options == null) options = undefined;
                    if (typeof (options) == 'string' || typeof (options) == 'number') {
                        if (options.toString().toLowerCase() == 'now') hideMS = 0;
                        if (parseInt(options) == options) hideMS = options;
                    }

                    setTimeout(function () {
                        $dialog.modal('hide');
                    }, hideMS);
                }
            };
        })(jQuery);
    }
    
    $.loadingDialog = waitModeling;
})();
