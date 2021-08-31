/* ====================================================================================================
/ Bootstrap Form Function: Form相關程式
/  開發人員:Tony Chen (20210204)
/ 範例: 執行: $("form#" + PageInfo.PROG_ID).clearForm();
/       
/ 參數:  {
/                   
/                   
/                }
/ ===================================================================================================== */

//Reset Validator 重設驗證
(function ($) {

    $.fn.clearForm = function (options) {

        // This is the easiest way to have default options.
        var settings = $.extend({
            // These are the defaults.

            formId: this.closest('form')

        }, options);

        var $form = $(settings.formId);

        //reset jQuery Validate's internals
        $form.validate().resetForm();

        //reset unobtrusive validation summary, if it exists
        $form.find("[data-valmsg-summary=true]")
            .removeClass("validation-summary-errors")
            .addClass("validation-summary-valid")
            .find("ul").empty();

        //reset unobtrusive field level, if it exists
        $form.find("[data-valmsg-replace]")
            .removeClass("field-validation-error")
            .addClass("field-validation-valid")
            .empty();


        return $form;
    };

}(jQuery));


//Reset Validator 重設驗證
function removeValidationErrors(frmId) {
    var $myform = $('#' + frmId);
    $myform.get(0).reset();
    var $myValidator = $myform.validate();
    $($myform).removeData('validator');
    $($myform).removeData('unobtrusiveValidation');
    $.validator.unobtrusive.parse($myform);
    $myValidator.resetForm();
    $('#' + frmId + ' input, select').removeClass('input-validation-error');
}

function removeValidationErrorsV2(frmId) {
    var $myform = $('#' + frmId);
    $myform.get(0).reset();
    var $myValidator = $myform.validate();
    $($myform).removeData('validator');
    $($myform).removeData('unobtrusiveValidation');
    $.validator.unobtrusive.parse($myform);
    $myValidator.resetForm();

    //reset unobtrusive validation summary, if it exists
    $myform.find("[data-valmsg-summary=true]")
        .removeClass("validation-summary-errors")
        .addClass("validation-summary-valid")
        .find("ul").empty();

    //reset unobtrusive field level, if it exists
    $myform.find("[data-valmsg-replace]")
        .removeClass("field-validation-error")
        .addClass("field-validation-valid")
        .removeData("unobtrusiveContainer")
        .find(">*")  // If we were using valmsg-replace, get the underlying error
        .removeData("unobtrusiveContainer")
        .empty();


    $('#' + frmId + ' input, select').removeClass('input-validation-error');
}



    /*
    (function ($) {
        //re-set all client validation given a jQuery selected form or child
        $.fn.resetValidation = function () {
            //var $form = this.closest('form');
            var $form = $(this);

            //reset jQuery Validate's internals
            $form.validate().resetForm();

            //reset unobtrusive validation summary, if it exists
            $form.find("[data-valmsg-summary=true]")
                .removeClass("validation-summary-errors")
                .addClass("validation-summary-valid")
                .find("ul").empty();

            //reset unobtrusive field level, if it exists
            $form.find("[data-valmsg-replace]")
                .removeClass("field-validation-error")
                .addClass("field-validation-valid")
                .empty();

            return $form;
        };
    })(jQuery);
    */