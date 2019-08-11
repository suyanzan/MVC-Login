using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary1
{
    public class Class1
    {

        public static void LockExcel(string filepath,int sheet_index,string password)
        {
            //Create workbook
            IWorkbook workBook;
            FileStream fsFile = new FileStream(filepath, FileMode.Open);
            workBook = new HSSFWorkbook(fsFile);
            string strFilePath = string.Format(filepath);
            using (FileStream fs = new FileStream(strFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                workBook = new HSSFWorkbook(fs);
            }
            //取得整份Excel之後，再去決定要去哪一個Sheet內資料。在NPOI每個Sheet都是一個陣列中的物件，故可以用Index去取
            //Protect the sheet
            HSSFSheet hst;
            hst = (HSSFSheet)workBook.GetSheetAt(sheet_index);
            //protect excel
            hst.ProtectSheet(password);
            //Save the file
            FileStream file = File.Create(filepath);
            workBook.Write(file);
            file.Close();

        }

        public static void UnLockExcel(string filepath, int sheet_index, string password)
        {
            //Create workbook
            IWorkbook workBook;
            FileStream fsFile = new FileStream(filepath, FileMode.Open);
            workBook = new HSSFWorkbook(fsFile);
            string strFilePath = string.Format(filepath);
            using (FileStream fs = new FileStream(strFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                workBook = new HSSFWorkbook(fs);
            }
            //取得整份Excel之後，再去決定要去哪一個Sheet內資料。在NPOI每個Sheet都是一個陣列中的物件，故可以用Index去取
            //Protect the sheet
            HSSFSheet hst;
            hst = (HSSFSheet)workBook.GetSheetAt(sheet_index);
            HSSFRow row1 = (HSSFRow)hst.CreateRow(0);
            HSSFCell cel1 = (HSSFCell)row1.CreateCell(0);
            HSSFCell cel2 = (HSSFCell)row1.CreateCell(1);
            ICellStyle unlocked = workBook.CreateCellStyle();
            unlocked.IsLocked = false;//設定為非鎖定
            //cel1.SetCellValue("未被锁定");
            cel1.CellStyle = unlocked;
            //Save the file
            FileStream file = File.Create(filepath);
            workBook.Write(file);
            file.Close();
            //sheet1.ProtectSheet("password");

        }
    }

}
