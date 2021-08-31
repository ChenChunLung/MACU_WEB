using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_TCC000.Controllers
{
    //從Remote DB或日記帳 讀取StoreInfo
    public class MERP_TCC002Controller : Controller
    {

        // GET: MERP_TCC000/MERP_TCC002
        public ActionResult Index()
        {
            return View();
        }

        // GET: MERP_TCC000/MERP_TCC002/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MERP_TCC000/MERP_TCC002/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: MERP_TCC000/MERP_TCC002/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: MERP_TCC000/MERP_TCC002/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MERP_TCC000/MERP_TCC002/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: MERP_TCC000/MERP_TCC002/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: MERP_TCC000/MERP_TCC002/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }


    }
}
