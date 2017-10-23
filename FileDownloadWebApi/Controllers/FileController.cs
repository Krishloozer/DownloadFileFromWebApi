using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Hosting;
using System.Web.Http;

namespace FileDownloadWebApi.Controllers
{
    [RoutePrefix("api/file")]
    public class FileController : ApiController
    {
        [HttpGet]
        [AllowAnonymous]
        [Route("download")]
        public IHttpActionResult Download() 
        {
            MemoryStream ms = new MemoryStream();
            var fileName = "Sample_" + DateTime.Now.ToString("ddMMyyyyhhmmss");
            var extension = "xls";
            var filepath = GenerateExcel(GetDataTable(), fileName, extension, out ms);

            //MemoryStream ms = new MemoryStream();
            //using (FileStream file = new FileStream(@"E:\Sample.xls", FileMode.Open, FileAccess.Read))
            //{
            //    file.CopyTo(ms);
            //}
            var result = new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(ms.GetBuffer()) };
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") { FileName = (fileName + "." + extension) };
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/ms-excel");
            var response = ResponseMessage(result);
            return response;
        }

        private string GenerateExcel(DataTable dataTable, string fileName, string extension, out MemoryStream mstemp) 
        {
            IWorkbook workbook;

            if (extension == "xlsx")
            {
                workbook = new XSSFWorkbook();
            }
            else if (extension == "xls")
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                throw new Exception("This format is not supported");
            }

            ISheet sheet1 = workbook.CreateSheet(fileName);

            //make a header row
            IRow row1 = sheet1.CreateRow(0);

            for (int j = 0; j < dataTable.Columns.Count; j++)
            {

                ICell cell = row1.CreateCell(j);
                String columnName = dataTable.Columns[j].ToString();
                cell.SetCellValue(columnName);
            }

            //loops through data
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                IRow row = sheet1.CreateRow(i + 1);
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {

                    ICell cell = row.CreateCell(j);
                    String columnName = dataTable.Columns[j].ToString();
                    cell.SetCellValue(dataTable.Rows[i][columnName].ToString());
                }
            }

            var dir = HostingEnvironment.MapPath("~/Content");
            var filePath = Path.Combine(dir, fileName+"."+extension);
            FileStream xfile = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            workbook.Write(xfile);
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            mstemp = ms;
            xfile.Close();
            return filePath;
        }

        private DataTable GetDataTable() 
        {
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            // Here we add five DataRows.
            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            return table;
        }
    }
}