using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.MCC
{
    public static class Extensions
    {
        public static string TemplateReplace(this string template, DataTable dtPrintData)
        {
            int count;
            string output = string.Empty;

            for (int i = 0; i < dtPrintData.Columns.Count; i++)
            {
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName + ">", dtPrintData.Rows[0][i].ToString());
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName.ToUpper() + ">", dtPrintData.Rows[0][i].ToString());
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName.ToLower() + ">", dtPrintData.Rows[0][i].ToString());
            }

            count = 1;
            foreach (DataRow dr in dtPrintData.Rows)
            {
                for (int iCol = 0; iCol < dtPrintData.Columns.Count; iCol++)
                {
                    //SN1~SNn
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToString() + count.ToString() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToUpper() + count.ToString() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToLower() + count.ToString() + ">", dr[iCol].ToString());

                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToString() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToUpper() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToLower() + ">", dr[iCol].ToString());
                }
                count++;
            }

            for (int iSeq = count; iSeq < 30; iSeq++)
            {
                for (int iCol = 0; iCol < dtPrintData.Columns.Count; iCol++)
                {
                    //SN1~SNn
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToString() + iSeq.ToString() + ">", "<DEL_LINE>");
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToUpper() + iSeq.ToString() + ">", "<DEL_LINE>");
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToLower() + iSeq.ToString() + ">", "<DEL_LINE>");
                }
            }

            string[] lines = template.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("<DEL_LINE>") == false)
                {
                    output += lines[i].ToString() + "\n";
                }
            }

            return output;
        }
    }
}
