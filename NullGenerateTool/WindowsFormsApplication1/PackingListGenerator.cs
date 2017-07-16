using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO; // File.Exists()

using NPOI.XSSF.UserModel; // XSSFWorkbook, XSSFSheet
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace NULL_is_my_son
{
    class PackingListGenerator
    {

        private String errorCode;
        private String file;

        private XSSFWorkbook wb;
        private XSSFSheet sheet_1;
        private XSSFSheet sheet_2;

        private Database databaseInfor;
        private ParkingListParser packingParser;

        private PackingListSheet1Generator sheet_1_generator;
        private PackingListDetailSheet2Generator sheet_2_generator;

        public PackingListGenerator(ParkingListParser packingParser, Database databaseInfor, String file)
        {
            this.databaseInfor = databaseInfor;
            this.packingParser = packingParser;
            this.file = file;

            this.errorCode = "";

            wb = new XSSFWorkbook();

            // create sheet
            sheet_1 = (XSSFSheet)wb.CreateSheet("工作表2");
            sheet_2 = (XSSFSheet)wb.CreateSheet("PL detail");


            sheet_1_generator = new PackingListSheet1Generator(wb, sheet_1, databaseInfor, packingParser);
            sheet_2_generator = new PackingListDetailSheet2Generator(wb, sheet_2, databaseInfor, packingParser);

        }

        public bool GeneratePackingListFile()
        {

            if (System.IO.File.Exists(file))
            {
                // Use a try block to catch IOExceptions, to
                // handle the case of the file already being
                // opened by another process.
                try
                {
                    System.IO.File.Delete(file);
                }
                catch (System.IO.IOException e)
                {
                    this.errorCode = "Delete file:" + e.Message;
                    return false;
                }
            }

            try
            {
                FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write);


                sheet_1_generator.GeneratePackingListSheet_1();
                sheet_2_generator.PackingListDetailGeneratorSheet2();

                wb.Write(fs);

                fs.Close();
            }
            catch (System.IO.IOException e)
            {
                this.errorCode = "Open file:" + e.Message;
                return false;
            }

            return true;
        }
    }
}
