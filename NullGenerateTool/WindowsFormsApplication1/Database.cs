using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO; // File.Exists()

using NPOI.XSSF.UserModel; // XSSFWorkbook, XSSFSheet

namespace NULL_is_my_son
{
    class Database
    {
        private List<DatabaseItem> listDatabase;
        private string errorCode;


        public Database()
        {
            this.listDatabase = null;
            this.errorCode = "";
        }

        public bool CreateListDatabase(String pathFile)
        {
            bool addItem;
            DatabaseItem item = null;

            if (!File.Exists(pathFile))
            {
                this.errorCode = "Find not found";
                return false;
            }

            listDatabase = new List<DatabaseItem>();

            try
            {

                // get sheets list from xlsx
                FileStream fs = new FileStream(pathFile, FileMode.Open, FileAccess.Read);

                XSSFWorkbook wb = new XSSFWorkbook(fs);

                if (wb.Count <= 0)
                {
                    this.errorCode = "Sheet in file is empty";
                    return false;
                }

                XSSFSheet sh = (XSSFSheet)wb.GetSheetAt(0);

                int index = 0;
                while (sh.GetRow(index) != null)
                {
                    if (index == 0)
                    {
                        index++;

                        continue;
                    }

                    addItem = false;


                    // write row value
                    for (int j = 0; j < sh.GetRow(index).Cells.Count; j++)
                    {
                        var cell = sh.GetRow(index).GetCell(j);

                        if (j == 0 && cell != null)
                        {
                            if (cell.CellType != NPOI.SS.UserModel.CellType.String)
                            {
                                break;
                            }
                            else if (sh.GetRow(index).GetCell(j).StringCellValue == null ||
                                sh.GetRow(index).GetCell(j).StringCellValue == "")
                            {
                                break;
                            }
                            else
                            {
                                item = new DatabaseItem();
                            }
                        }

                        if (cell != null)
                        {
                            // TODO: you can add more cell types capatibility, e. g. formula
                            switch (cell.CellType)
                            {
                                case NPOI.SS.UserModel.CellType.Numeric:
                                    
                                    addItem = true;
                                    item.SetDataValue(j, sh.GetRow(index).GetCell(j).NumericCellValue + "");
                                    break;
                                case NPOI.SS.UserModel.CellType.String:
                                    addItem = true;
                                    item.SetDataValue(j, sh.GetRow(index).GetCell(j).StringCellValue + "");
                                    
                                    break;
                            }
                        }
                    }
                    if (addItem)
                    {
                        listDatabase.Add(item);
                    }
                    index++;
                }

                fs.Close();
            }
            catch (IOException ex)
            {
                this.errorCode = ex.Message;

                return false;
            }

            return true;
        }


        public string GetErrorCode()
        {
            return this.errorCode;
        }

        public DatabaseItem FindItem(string nameProduct)
        {
            if (this.listDatabase == null)
            {
                this.errorCode = "Database is null";
                return null;
            }

            if (this.listDatabase.Count <= 0)
            {
                this.errorCode = "Database is empty";
                return null;
            }

            foreach (DatabaseItem item in this.listDatabase) // Display for verification.
            {
                if (item.nameProduct == nameProduct)
                {
                    return item;
                }
            }

            this.errorCode = "Not found";
            return null;
        }

    }
}
