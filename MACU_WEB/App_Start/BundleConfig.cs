using System.Web;
using System.Web.Optimization;

namespace MACU_WEB
{
    public class BundleConfig
    {
        // 如需統合的詳細資訊，請瀏覽 https://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-{version}.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Scripts/jquery.validate*"));

            // 使用開發版本的 Modernizr 進行開發並學習。然後，當您
            // 準備好可進行生產時，請使用 https://modernizr.com 的建置工具，只挑選您需要的測試。
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                      "~/Scripts/bootstrap.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/bootstrap.css",
                      "~/Content/site.css"));

            //20201216 CCL+ Str 用於JsFunLoading///////////////////////////////////////////////////////
            //放系統的JS Scripts (For Ajax), PS:安裝unobtrusive-ajax NuGet套件系統不會自動加入
            //要自己手動加
            bundles.Add(new ScriptBundle("~/bundles/SysScripts").Include(
                    "~/Scripts/jquery.unobtrusive-ajax.min.js"
                    
                ));

            //Common放自己寫的外掛 --------------------------------------------------
            bundles.Add(new ScriptBundle("~/bundles/Common").Include(
                    "~/Scripts/Common/JS_UI.js"
                    , "~/Scripts/Common/JS_UI_MESSAGE.js"
                    , "~/Scripts/Common/JS_FORMS.js"
                ));

            //Common放自己寫的CSS 
            bundles.Add(new StyleBundle("~/Content/Common").Include(
                    "~/Content/Common/macu-style.css"
                ));

            
            //Plugins 放網上下載的外掛 ---------------------------------------------
            bundles.Add(new ScriptBundle("~/bundles/Plugins").Include(
                    "~/Scripts/Plugins/bootstrap-dialog/bootstrap-dialog.min.js"
                    , "~/Scripts/Plugins/jquery-validation/localization/messages_zh_TW.min.js"
                    , "~/Scripts/Plugins/jquery-validation/additional-methods.min.js"
                    , "~/Scripts/Plugins/jquery-ui/jquery-ui.min.js"
                    , "~/Scripts/Plugins/jqgrid/i18n/grid.locale-tw.js"
                    , "~/Scripts/Plugins/jqgrid/jquery.jqGrid.min.js"
                    , "~/Scripts/Plugins/jqgrid/plugins/*.js"
                ));

            bundles.Add(new StyleBundle("~/Content/Plugins").Include(
                      "~/Content/Plugins/bootstrap-dialog/bootstrap-dialog.min.css"
                      , "~/Content/Plugins/jquery-validation/jquery-validation.css"
                      //改用bootstrap專用 css , "~/Content/Plugins/jqGrid/ui.jqgrid.css"
                      , "~/Content/Plugins/jquery-ui/jquery-ui.min.css"
                      , "~/Content/Plugins/jqGrid/ui.jqgrid-bootstrap.css"
                      , "~/Content/Plugins/jqGrid/ui.jqgrid-bootstrap-ui.css"
                      , "~/Content/Plugins/jqGrid/plugins/*.css"
               ));

            //Theme 風格化 --------------------------------------------------------           
            bundles.Add(new StyleBundle("~/Content/Theme-Base/css").Include(
                      "~/Content/Theme/assets/css/*.css",
                      "~/Content/Theme/assets/font-awesome/css/*.css",
                      "~/Content/Theme/assets/js/gritter/css/jquery.gritter.css",
                      "~/Content/Theme/assets/lineicons/style.css"                    
                      ));

            bundles.Add(new StyleBundle("~/Content/Theme-Form/css").Include(                   
                      "~/Content/Theme/assets/js/bootstrap-datepicker/css/datepicker.css",
                      "~/Content/Theme/assets/js/bootstrap-daterangepicker/daterangepicker.css"
                      ));

            bundles.Add(new StyleBundle("~/Content/Theme-Table/css").Include(
                     "~/Content/Theme/assets/css/style-responsive.css",
                     "~/Content/Theme/assets/css/table-responsive.css"
                      ));

            bundles.Add(new ScriptBundle("~/Content/Theme-Base/js").Include(
                      "~/Content/Theme/assets/js/chart-master/Chart.js",
                      "~/Content/Theme/assets/js/jquery.dcjqaccordion.2.7.js",
                      "~/Content/Theme/assets/js/jquery.scrollTo.min.js",
                      "~/Content/Theme/assets/js/jquery.nicescroll.js",
                      "~/Content/Theme/assets/js/jquery.sparkline.js",
                      "~/Content/Theme/assets/js/common-scripts.js",
                      "~/Content/Theme/assets/js/gritter/js/jquery.gritter.js",
                      "~/Content/Theme/assets/js/gritter-conf.js",
                      "~/Content/Theme/assets/js/sparkline-chart.js",
                      "~/Content/Theme/assets/js/zabuto_calendar.js"                                        
                      ));
            //

            bundles.Add(new ScriptBundle("~/Content/Theme-Form/js").Include(
                      "~/Content/Theme/assets/js/jquery-ui-1.9.2.custom.min.js",
                      "~/Content/Theme/assets/js/bootstrap-switch.js",
                      "~/Content/Theme/assets/js/jquery.tagsinput.js",
                      "~/Content/Theme/assets/js/bootstrap-datepicker/js/bootstrap-datepicker.js",
                      "~/Content/Theme/assets/js/bootstrap-daterangepicker/date.js",
                      "~/Content/Theme/assets/js/bootstrap-daterangepicker/daterangepicker.js",
                      "~/Content/Theme/assets/js/bootstrap-inputmask/bootstrap-inputmask.min.js",
                      "~/Content/Theme/assets/js/form-component.js"
                      ));

            //20201216 CCL+ End 用於JsFunLoading///////////////////////////////////////////////////////

        }
    }
}
