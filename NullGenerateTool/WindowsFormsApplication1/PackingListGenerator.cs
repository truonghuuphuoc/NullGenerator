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
        private String[] Sheet_1_Header = {   "日期", "貨號",     "顏色", "顏色2", "尺寸", 
                                              "QTY",  "箱數",     "箱號", "合計",  "淨重", 
                                              "毛重", "外箱規格", "備註"};

        private List<String> Sheet_2_size_1 = null;

        private String[] Sheet_2_size_2 = { "M",   "L",   "XL",   "XXL"};

        private String errorCode;
        private String file;

        private XSSFWorkbook wb;
        private XSSFSheet sheet_1;
        private XSSFSheet sheet_2;

        private Database databaseInfor;
        private ParkingListParser packingParser;

        private int rowIndex_sheet_1;
        private int sheet_2_rowIndex;

        private double numberPacket_sheet_1;
        private double sheet_2_numberPacket;

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

            rowIndex_sheet_1 = 0;
            sheet_2_rowIndex = 0;

            var r = sheet_1.CreateRow(rowIndex_sheet_1);
            rowIndex_sheet_1++;

            for (int idex = 0; idex < Sheet_1_Header.Length; idex++)
            {
                r.CreateCell(idex).SetCellValue(Sheet_1_Header[idex]);
            }
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

                if (packingParser.GetParkingParserList() != null
                    && packingParser.GetParkingParserList().Count > 0)
                {
                    numberPacket_sheet_1 = 0;
                    sheet_2_numberPacket = 0;

                    foreach (ParkingParserItem item in packingParser.GetParkingParserList())
                    {
                        if (item.GetListItem() != null && item.GetListItem().Count > 0)
                        {
                            foreach (PackingListItem baseEle in item.GetListItem())
                            {
                                if(CheckExistsSize_2(baseEle.GetProductSize()) < 0)
                                {
                                    AddToListSize_1(baseEle.GetProductSize());
                                }

                                if (baseEle.GetMergerList() != null && baseEle.GetMergerList().Count > 0)
                                {
                                    foreach (MergerItem merger in baseEle.GetMergerList())
                                    {
                                        if (CheckExistsSize_2(merger.GetProductSize()) < 0)
                                        {
                                            AddToListSize_1(merger.GetProductSize());
                                        }
                                    }
                                }
                            }
                        }
                    }

                    foreach (ParkingParserItem item in packingParser.GetParkingParserList())
                    {
                        if (item.GetListItem() != null && item.GetListItem().Count > 0)
                        {
                            foreach (PackingListItem baseEle in item.GetListItem())
                            {
                                CreateRowValueSheet_1(baseEle, item.GetDatabaseItem());
                            }
                        }
                    }

                    foreach (ParkingParserItem item in packingParser.GetParkingParserList())
                    {
                        CreateRowValueSheet_2(item);
                    }
                }


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

        ////////////////////////////////////////////////////////////////////
        private void CreateRowValueSheet_1(PackingListItem item, DatabaseItem databaseItem)
        {
            NPOI.SS.UserModel.IRow r;

            if (item.GetQuantity() < databaseItem.GetMaxPacketSize())
            {
                r = sheet_1.CreateRow(rowIndex_sheet_1);
                rowIndex_sheet_1++;

                numberPacket_sheet_1++;

                double valueMerger = item.GetQuantity();

                if (item.GetMergerList() != null && item.GetMergerList().Count > 0)
                {
                    foreach (MergerItem merger in item.GetMergerList())
                    {
                        valueMerger += merger.GetQuantity();
                    }
                }

                for (int idex = 0; idex < Sheet_1_Header.Length; idex++)
                {
                    switch (idex)
                    {
                        case 0:
                            r.CreateCell(idex).SetCellValue(this.packingParser.GetDateTime());
                            break;

                        case 1:
                            r.CreateCell(idex).SetCellValue(item.GetNameProduct());
                            break;

                        case 2:
                            r.CreateCell(idex).SetCellValue(item.GetColor1());
                            break;

                        case 3:
                            r.CreateCell(idex).SetCellValue(item.GetColor2());
                            break;

                        case 4:
                            r.CreateCell(idex).SetCellValue(item.GetProductSize());
                            break;

                        case 5:
                            {
                                ICell cell = r.CreateCell(idex);
                                cell.SetCellValue(item.GetQuantity());

                                ICellStyle style = wb.CreateCellStyle();
                                style.FillForegroundColor = IndexedColors.Yellow.Index;
                                style.FillPattern = FillPattern.SolidForeground;

                                cell.CellStyle = style; 
                                break;
                            }

                        case 6:
                            {
                                ICell cell = r.CreateCell(idex);
                                if (valueMerger > item.GetQuantity())
                                {
                                    ICellStyle style = wb.CreateCellStyle();
                                    style.VerticalAlignment = VerticalAlignment.Center;
                                    cell.CellStyle = style;
                                }
                                cell.SetCellValue(1);
                                break;
                            }

                        case 7:
                            r.CreateCell(idex).SetCellValue(numberPacket_sheet_1);
                            break;

                        case 8:
                            {
                                ICell cell = r.CreateCell(idex);
                                if (valueMerger > item.GetQuantity())
                                {
                                    ICellStyle style = wb.CreateCellStyle();
                                    style.VerticalAlignment = VerticalAlignment.Center;
                                    cell.CellStyle = style;
                                }
                                cell.SetCellValue(valueMerger);
                                break;
                            }

                        case 9:
                            r.CreateCell(idex).SetCellValue(databaseItem.GetNetWeight());
                            break;

                        case 10:
                            r.CreateCell(idex).SetCellValue(databaseItem.GetAllWeight());
                            break;

                        case 11:
                            r.CreateCell(idex).SetCellValue(databaseItem.GetPacketInformation());
                            break;

                        default:
                            break;
                    }
                }
            }
            else
            {
                for(int tmp = 0; tmp < (int)(item.GetQuantity()/ databaseItem.GetMaxPacketSize()); tmp ++)
                {
                    numberPacket_sheet_1++;

                    r = sheet_1.CreateRow(rowIndex_sheet_1);
                    rowIndex_sheet_1++;

                    for (int idex = 0; idex < Sheet_1_Header.Length; idex++)
                    {
                        switch (idex)
                        {
                            case 0:
                                r.CreateCell(idex).SetCellValue(this.packingParser.GetDateTime());
                                break;

                            case 1:
                                r.CreateCell(idex).SetCellValue(item.GetNameProduct());
                                break;

                            case 2:
                                r.CreateCell(idex).SetCellValue(item.GetColor1());
                                break;

                            case 3:
                                r.CreateCell(idex).SetCellValue(item.GetColor2());
                                break;

                            case 4:
                                r.CreateCell(idex).SetCellValue(item.GetProductSize());
                                break;

                            case 5:
                                r.CreateCell(idex).SetCellValue(databaseItem.GetMaxPacketSize());
                                break;

                            case 6:
                                r.CreateCell(idex).SetCellValue(1);
                                break;

                            case 7:
                                r.CreateCell(idex).SetCellValue(numberPacket_sheet_1);
                                break;

                            case 8:
                                r.CreateCell(idex).SetCellValue(databaseItem.GetMaxPacketSize());
                                break;

                            case 9:
                                r.CreateCell(idex).SetCellValue(databaseItem.GetNetWeight());
                                break;

                            case 10:
                                r.CreateCell(idex).SetCellValue(databaseItem.GetAllWeight());
                                break;

                            case 11:
                                r.CreateCell(idex).SetCellValue(databaseItem.GetPacketInformation());
                                break;

                            default:
                                break;
                        }
                    }
                }
            }

            if (item.GetMergerList() != null && item.GetMergerList().Count > 0)
            {
                foreach (MergerItem merger in item.GetMergerList())
                {
                    r = sheet_1.CreateRow(rowIndex_sheet_1);
                    rowIndex_sheet_1++;

                    for (int idex = 0; idex < Sheet_1_Header.Length; idex++)
                    {
                        switch (idex)
                        {
                            case 0:
                                r.CreateCell(idex).SetCellValue(this.packingParser.GetDateTime());
                                break;

                            case 1:
                                r.CreateCell(idex).SetCellValue(merger.GetNameProduct());
                                break;

                            case 2:
                                r.CreateCell(idex).SetCellValue(merger.GetColor1());
                                break;

                            case 3:
                                r.CreateCell(idex).SetCellValue(merger.GetColor2());
                                break;

                            case 4:
                                r.CreateCell(idex).SetCellValue(merger.GetProductSize());
                                break;

                            case 5:
                                {
                                    ICell cell = r.CreateCell(idex);
                                    cell.SetCellValue(merger.GetQuantity());
                                    ICellStyle style = wb.CreateCellStyle();
                                    style.FillForegroundColor = IndexedColors.Yellow.Index;
                                    style.FillPattern = FillPattern.SolidForeground;

                                    cell.CellStyle = style; 
                                    break;
                                }

                            case 6:
                                r.CreateCell(idex);
                                break;

                            case 7:
                                r.CreateCell(idex).SetCellValue(numberPacket_sheet_1);
                                break;

                            case 8:
                                r.CreateCell(idex);
                                break;

                            case 9:
                                r.CreateCell(idex).SetCellValue(databaseItem.GetNetWeight());
                                break;

                            case 10:
                                r.CreateCell(idex).SetCellValue(databaseItem.GetAllWeight());
                                break;

                            case 11:
                                r.CreateCell(idex).SetCellValue(databaseItem.GetPacketInformation());
                                break;

                            default:
                                break;
                        }
                    }
                }

                NPOI.SS.Util.CellRangeAddress range_1 = new NPOI.SS.Util.CellRangeAddress(rowIndex_sheet_1 -( 1 + item.GetMergerList().Count),
                                                                                          rowIndex_sheet_1 - 1, 6, 6);

                NPOI.SS.Util.CellRangeAddress range_2 = new NPOI.SS.Util.CellRangeAddress(rowIndex_sheet_1 - (1 + item.GetMergerList().Count),
                                                                                          rowIndex_sheet_1 - 1, 8, 8);
                sheet_1.AddMergedRegion(range_1);
                sheet_1.AddMergedRegion(range_2);
                       
            }
        }


        private void CreateRowValueSheet_2(ParkingParserItem item)
        {
            NPOI.SS.UserModel.IRow row;
            ICell cell;
            int type;

            if(!(item != null && item.GetListItem() != null && item.GetListItem().Count > 0))
            {
                return;
            }

            //Clear header value
            row = sheet_2.CreateRow(sheet_2_rowIndex);
            sheet_2_rowIndex++;

            if(CheckExistsSize_2(item.GetListItem().ElementAt(0).GetProductSize()) < 0)
            {
                type = 1;
            }
            else
            {
                type = 2;
            }

            for (int run = 0; run < Sheet_2_GetNumberOfColum(); run++)
            {
                cell = row.CreateCell(run);

                cell.SetCellValue(Sheet_2_GetHeaderString(type, run));

                if(run == 0)
                {
                    cell.SetCellValue(item.GetNameProduct());
                }
            }

            NPOI.SS.Util.CellRangeAddress range = new NPOI.SS.Util.CellRangeAddress(sheet_2_rowIndex - 1, sheet_2_rowIndex - 1, 0, 4);                                                                           
            sheet_2.AddMergedRegion(range);

            foreach(PackingListItem packItem in item.GetListItem())
            {
                //create new row
                row = sheet_2.CreateRow(sheet_2_rowIndex);
                sheet_2_rowIndex++;


                for (int run = 0; run < Sheet_2_GetNumberOfColum(); run++)
                {
                }

                if (packItem.GetMergerList() != null && packItem.GetMergerList().Count > 0)
                {
                    //insert to same row
                }
            }


        }

        private int AddToListSize_1(String size)
        {
            int index = -1;

            if (Sheet_2_size_1 == null)
            {
                Sheet_2_size_1 = new List<string>();
            }

            for (int run = 0; run < Sheet_2_size_1.Count ; run++)
            {
                if(size == Sheet_2_size_1.ElementAt(run))
                {
                    return run;
                }

                if (String.Compare(size, Sheet_2_size_1.ElementAt(run)) > 0)
                {
                    index = run;
                }
            }

            if (index == -1)
            {
                Sheet_2_size_1.Add(size);
            }
            else
            {
                Sheet_2_size_1.Insert(index + 1, size);
            }
            return -1;
        }

        private int CheckExistsSize_2(String size)
        {
            for (int run = 0; run < Sheet_2_size_2.Length; run++)
            {
                if (size == Sheet_2_size_2[run])
                {
                    return run;
                }
            }
            return -1;
        }

        private int Sheet_2_GetNumberOfColum()
        {
            if(Sheet_2_size_1 == null)
            {
                Sheet_2_size_1 = new List<string>();
            }

            int value = Sheet_2_size_2.Length > Sheet_2_size_1.Count ? Sheet_2_size_1.Count : Sheet_2_size_2.Length;

            value += 5;

            return value;
        }

        private String Sheet_2_GetHeaderString(int type, int index)
        {
            if (index < 5)
            {
                return "";
            }
            else
            {
                index -= 5;
            }

            if (type == 1)
            {
                if (index < Sheet_2_size_1.Count)
                {
                    return Sheet_2_size_1.ElementAt(index);
                }
                else
                {
                    return "";
                }
            }
            else
            {
                if (index < Sheet_2_size_2.Length)
                {
                    return Sheet_2_size_2[index];
                }
                else
                {
                    return "";
                }
            }

            return "";
            
        }

        private int Sheet_2_GetColumIndex(int type, String  size)
        {

            int index = 0;

            if (type == 1)
            {
                for (int run = 0; run < Sheet_2_size_1.Count; run++)
                {
                    if (size == Sheet_2_size_1.ElementAt(run))
                    {
                        index = run;
                        break;
                    }
                }
            }
            else
            {
                for (int run = 0; run < Sheet_2_size_2.Length; run++)
                {
                    if (size == Sheet_2_size_2[run])
                    {
                        index = run;
                        break;
                    }
                }
            }

            return  (index + 5);

        }
    }
}
