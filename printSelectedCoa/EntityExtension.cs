using System;
using System.Collections.Generic;
using System.Data.Metadata.Edm;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using One1.Controls;


namespace printSelectedCoa
{


    [ComVisible(true)]
    [ProgId("printSelectedCoa.printSelectedCoacls")]
    public class printSelectedCoa : IEntityExtension
    {
        private INautilusServiceProvider sp;

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }

        public void Execute(ref LSExtensionParameters Parameters)
        {
            sp = Parameters["SERVICE_PROVIDER"];
            var ntlsCon = Utils.GetNtlsCon(sp);
            Utils.CreateConstring(ntlsCon);

            var dal = new DataLayer();
            dal.Connect();
            var records = Parameters["RECORDS"];

            List<string> Ids = new List<string>();
            while (!records.EOF)
            {
                var id = records.Fields["U_COA_REPORT_ID"].Value;
                Ids.Add(id);
                records.MoveNext();
            }
            foreach (var id in Ids)
            {
                if (id != null)
                {
                    var coa = dal.GetCoaReportById(long.Parse(id));
                    string documentToPrint;

                    if (coa.PdfPath != null)
                    {
                        documentToPrint = coa.PdfPath;
                    }
                    else
                    {
                        documentToPrint = coa.DocPath;
                    }
                    if (documentToPrint == null)
                    {
                        CustomMessageBox.Show("Path doesn't exists", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        continue;
                    }

                    Process printJob = new Process();
                    printJob.StartInfo.FileName = documentToPrint;
                    printJob.StartInfo.UseShellExecute = true;
                    printJob.StartInfo.Verb = "print";

                    Thread oThread = new Thread(() => printJob.Start());
                    oThread.Start();


                }
            }



        }



    }
}
