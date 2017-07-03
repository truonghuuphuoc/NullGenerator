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

        private ICellStyle sheet_1_headerStyle_2;
        private ICellStyle sheet_1_headerStyle;

        private ICellStyle sheet_2_headerStyle_1;
        private ICellStyle sheet_2_headerStyle_2;

        private ICellStyle sheet_2_bottomStyle_1;
        private ICellStyle sheet_2_bottomStyle_2;
        private ICellStyle sheet_2_mainStyle_2;

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

            sheet_1_headerStyle = wb.CreateCellStyle();

            sheet_1_headerStyle.Alignment = HorizontalAlignment.Center;
            sheet_1_headerStyle.VerticalAlignment = VerticalAlignment.Center;
            sheet_1_headerStyle.BorderTop = BorderStyle.Thin;
            sheet_1_headerStyle.BorderRight = BorderStyle.Thin;
            sheet_1_headerStyle.BorderBottom = BorderStyle.Thin;
            sheet_1_headerStyle.BorderLeft = BorderStyle.Thin;


            sheet_1_headerStyle_2 = wb.CreateCellStyle();

            sheet_1_headerStyle_2.Alignment = HorizontalAlignment.Center;
            sheet_1_headerStyle_2.VerticalAlignment = VerticalAlignment.Center;
            sheet_1_headerStyle_2.BorderTop = BorderStyle.Thin;
            sheet_1_headerStyle_2.BorderRight = BorderStyle.Thin;
            sheet_1_headerStyle_2.BorderBottom = BorderStyle.Thin;
            sheet_1_headerStyle_2.BorderLeft = BorderStyle.Thin;
            sheet_1_headerStyle_2.FillForegroundColor = IndexedColors.Yellow.Index;
            sheet_1_headerStyle_2.FillPattern = FillPattern.SolidForeground;


            sheet_2_headerStyle_1 = wb.CreateCellStyle();
            IFont headerbold_1 = wb.CreateFont();
            headerbold_1.Boldweight = (short)FontBoldWeight.Bold;
            headerbold_1.FontHeight = 13;

            sheet_2_headerStyle_1.SetFont(headerbold_1);
            sheet_2_headerStyle_1.Alignment = HorizontalAlignment.Center;
            sheet_2_headerStyle_1.VerticalAlignment = VerticalAlignment.Center;
            sheet_2_headerStyle_1.BorderTop = BorderStyle.Thin;
            sheet_2_headerStyle_1.BorderRight = BorderStyle.Thin;
            sheet_2_headerStyle_1.BorderBottom = BorderStyle.Thin;

            sheet_2_headerStyle_2 = wb.CreateCellStyle();
            IFont headerbold_2 = wb.CreateFont();
            headerbold_2.Boldweight = (short)FontBoldWeight.Bold;
            headerbold_2.FontHeight = 12;

            headerbold_2.Underline = FontUnderlineType.Single;

            sheet_2_headerStyle_2.SetFont(headerbold_2);
            sheet_2_headerStyle_2.Alignment = HorizontalAlignment.Center;
            sheet_2_headerStyle_2.VerticalAlignment = VerticalAlignment.Center;
            sheet_2_headerStyle_2.BorderTop = BorderStyle.Thin;
            sheet_2_headerStyle_2.BorderRight = BorderStyle.Thin;
            sheet_2_headerStyle_2.BorderBottom = BorderStyle.Thin;

            sheet_2_bottomStyle_1 = wb.CreateCellStyle();
            sheet_2_bottomStyle_1.FillForegroundColor = IndexedColors.Aqua.Index;
            sheet_2_bottomStyle_1.FillPattern = FillPattern.SolidForeground;

            sheet_2_bottomStyle_1.BorderBottom = BorderStyle.Thin;
            sheet_2_bottomStyle_1.BorderLeft = BorderStyle.Thin;
            sheet_2_bottomStyle_1.BorderRight = BorderStyle.Thin;

            IFont boldFont_1 = wb.CreateFont();
            boldFont_1.Boldweight = (short)FontBoldWeight.Bold;
            boldFont_1.FontHeight = 13;
            boldFont_1.Color = IndexedColors.Red.Index;
            sheet_2_bottomStyle_1.SetFont(boldFont_1);

            sheet_2_bottomStyle_2 = wb.CreateCellStyle();
            sheet_2_bottomStyle_2.FillForegroundColor = IndexedColors.Aqua.Index;
            sheet_2_bottomStyle_2.FillPattern = FillPattern.SolidForeground;

            sheet_2_bottomStyle_2.BorderBottom = BorderStyle.Thin;
            sheet_2_bottomStyle_2.BorderLeft = BorderStyle.Thin;
            sheet_2_bottomStyle_2.BorderRight = BorderStyle.Thin;

            IFont boldFont_2 = wb.CreateFont();
            boldFont_2.Boldweight = (short)FontBoldWeight.Bold;
            boldFont_2.FontHeight = 13;
            boldFont_2.Color = IndexedColors.Black.Index;
            sheet_2_bottomStyle_2.SetFont(boldFont_2);


            sheet_2_mainStyle_2 = wb.CreateCellStyle();
            sheet_2_mainStyle_2.BorderBottom = BorderStyle.Dotted;
            sheet_2_mainStyle_2.BorderRight = BorderStyle.Thin;

            rowIndex_sheet_1 = 0;
            sheet_2_rowIndex = 0;

            var r = sheet_1.CreateRow(rowIndex_sheet_1);
            rowIndex_sheet_1++;

            for (int idex = 0; idex < Sheet_1_Header.Length; idex++)
            {
                var cell = r.CreateCell(idex);
                cell.SetCellValue(Sheet_1_Header[idex]);
                cell.CellStyle = sheet_1_headerStyle;
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
                    ICell cell = r.CreateCell(idex);
                    cell.CellStyle = sheet_1_headerStyle;
                    switch (idex)
                    {
                        case 0:
                            cell.SetCellValue(this.packingParser.GetDateTime());
                            break;

                        case 1:
                            cell.SetCellValue(item.GetNameProduct());
                            break;

                        case 2:
                            cell.SetCellValue(item.GetColor1());
                            break;

                        case 3:
                            cell.SetCellValue(item.GetColor2());
                            break;

                        case 4:
                            cell.SetCellValue(item.GetProductSize());
                            break;

                        case 5:
                            cell.SetCellValue(item.GetQuantity());
                            cell.CellStyle = sheet_1_headerStyle_2;
                            break;

                        case 6:
                            cell.SetCellValue(1);
                            break;

                        case 7:
                            cell.SetCellValue(numberPacket_sheet_1);
                            break;

                        case 8:
                            cell.SetCellValue(valueMerger);
                            break;

                        case 9:
                            cell.SetCellValue(databaseItem.GetNetWeight());
                            break;

                        case 10:
                            cell.SetCellValue(databaseItem.GetAllWeight());
                            break;

                        case 11:
                            cell.SetCellValue(databaseItem.GetPacketInformation());
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
                        ICell cell = r.CreateCell(idex);
                        cell.CellStyle = sheet_1_headerStyle;

                        switch (idex)
                        {
                            case 0:
                                cell.SetCellValue(this.packingParser.GetDateTime());
                                break;

                            case 1:
                                cell.SetCellValue(item.GetNameProduct());
                                break;

                            case 2:
                                cell.SetCellValue(item.GetColor1());
                                break;

                            case 3:
                                cell.SetCellValue(item.GetColor2());
                                break;

                            case 4:
                                cell.SetCellValue(item.GetProductSize());
                                break;

                            case 5:
                                cell.SetCellValue(databaseItem.GetMaxPacketSize());
                                break;

                            case 6:
                                cell.SetCellValue(1);
                                break;

                            case 7:
                                cell.SetCellValue(numberPacket_sheet_1);
                                break;

                            case 8:
                                cell.SetCellValue(databaseItem.GetMaxPacketSize());
                                break;

                            case 9:
                                cell.SetCellValue(databaseItem.GetNetWeight());
                                break;

                            case 10:
                                cell.SetCellValue(databaseItem.GetAllWeight());
                                break;

                            case 11:
                                cell.SetCellValue(databaseItem.GetPacketInformation());
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
                        ICell cell = r.CreateCell(idex);
                        cell.CellStyle = sheet_1_headerStyle;

                        switch (idex)
                        {
                            case 0:
                                cell.SetCellValue(this.packingParser.GetDateTime());
                                break;

                            case 1:
                                cell.SetCellValue(merger.GetNameProduct());
                                break;

                            case 2:
                                cell.SetCellValue(merger.GetColor1());
                                break;

                            case 3:
                                cell.SetCellValue(merger.GetColor2());
                                break;

                            case 4:
                                cell.SetCellValue(merger.GetProductSize());
                                break;

                            case 5:
                                cell.SetCellValue(merger.GetQuantity());
                                cell.CellStyle = sheet_1_headerStyle_2;
                                break;

                            case 6:
                                break;

                            case 7:
                                cell.SetCellValue(numberPacket_sheet_1);
                                break;

                            case 8:
                                break;

                            case 9:
                                cell.SetCellValue(databaseItem.GetNetWeight());
                                break;

                            case 10:
                                cell.SetCellValue(databaseItem.GetAllWeight());
                                break;

                            case 11:
                                cell.SetCellValue(databaseItem.GetPacketInformation());
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

            int start_row, end_row;

            if(!(item != null && item.GetListItem() != null && item.GetListItem().Count > 0))
            {
                return;
            }

            //Clear header value
            row = sheet_2.CreateRow(sheet_2_rowIndex);
            sheet_2_rowIndex++;
            start_row = sheet_2_rowIndex + 1;

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

                if (run == 0)
                {
                    cell.SetCellValue(item.GetNameProduct() + " - BRASSIERES");


                    cell.CellStyle = sheet_2_headerStyle_1;
                }
                else
                {
                    cell.CellStyle = sheet_2_headerStyle_2;
                }
            }

            NPOI.SS.Util.CellRangeAddress range = new NPOI.SS.Util.CellRangeAddress(sheet_2_rowIndex - 1, sheet_2_rowIndex - 1, 0, 4);                                                                           
            sheet_2.AddMergedRegion(range);


            foreach(PackingListItem packItem in item.GetListItem())
            {

                //create new row
                row = sheet_2.CreateRow(sheet_2_rowIndex);
                sheet_2_rowIndex++;


                int phan_nguyen = (int)(packItem.GetQuantity() / item.GetDatabaseItem().GetMaxPacketSize());
                int size_index = Sheet_2_GetColumIndex(type, packItem.GetProductSize());

                for (int run = 0; run < Sheet_2_GetNumberOfColum(); run++)
                {
                    cell = row.CreateCell(run);

                    cell.CellStyle = sheet_2_mainStyle_2;

                    switch (run)
                    {
                        case 0:
                            if (phan_nguyen > 1)
                            {
                                double from = sheet_2_numberPacket + 1;
                                double to = sheet_2_numberPacket + phan_nguyen;
                                cell.SetCellValue(from + " - " + to);

                                sheet_2_numberPacket += phan_nguyen;
                            }
                            else
                            {
                                sheet_2_numberPacket++;
                                cell.SetCellValue(sheet_2_numberPacket + "");
                            }
                            break;

                        case 1:
                            if (phan_nguyen > 0)
                            {
                                cell.SetCellValue(phan_nguyen);
                            }
                            else
                            {
                                cell.SetCellValue(1);
                            }
                            break;

                        case 2:
                            cell.SetCellValue(packItem.GetNameProduct());
                            break;

                        case 3:
                            cell.SetCellValue(packItem.GetColor1());
                            break;

                        case 4:
                            cell.SetCellValue(packItem.GetColor2());
                            break;

                        default:
                            if (run == size_index)
                            {
                                if (phan_nguyen > 0)
                                {
                                    String formula = item.GetDatabaseItem().GetMaxPacketSize() + "*$B" + sheet_2_rowIndex;
                                    cell.SetCellFormula(formula);
                                }
                                else
                                {
                                    cell.SetCellValue(packItem.GetQuantity());
                                }
                            }
                            else
                            {
                                int total = Sheet_2_GetNumberOfColum();

                                if((total - 4) == run)
                                {
                                    String columName = Sheet_2_GetStringOfColum(Sheet_2_GetNumberOfColum() - 4);
                                    String formula = "SUM(F" + sheet_2_rowIndex + ":" + columName + sheet_2_rowIndex + ")";
                                    cell.SetCellFormula(formula);
                                }
                                else  if((total - 3) == run)
                                {
                                    cell.SetCellValue("PCS");
                                }
                                else  if((total - 2) == run)
                                {
                                    String formula = item.GetDatabaseItem().GetNetWeight() + "*B" + sheet_2_rowIndex;
                                    cell.SetCellFormula(formula);
                                }
                                else  if((total - 1) == run)
                                {
                                    String formula = item.GetDatabaseItem().GetAllWeight() + "*B" + sheet_2_rowIndex;
                                    cell.SetCellFormula(formula);
                                }
                                else
                                {
                                    cell.SetCellValue("");
                                }
                               
                            }
                            break;
                    }
                }

                if (packItem.GetMergerList() != null && packItem.GetMergerList().Count > 0)
                {
                    //insert to same row
                    foreach (MergerItem mergerItem in packItem.GetMergerList())
                    {
                        size_index = Sheet_2_GetColumIndex(type, mergerItem.GetProductSize());

                        cell = row.GetCell(size_index);
                        cell.SetCellValue(mergerItem.GetQuantity());
                    }
                }
            }

            end_row = sheet_2_rowIndex;

            //sum row
            row = sheet_2.CreateRow(sheet_2_rowIndex);
            sheet_2_rowIndex++;

            for (int run = 0; run < Sheet_2_GetNumberOfColum(); run++)
            {
                cell = row.CreateCell(run);
                if (run == 2)
                {
                    cell.CellStyle = sheet_2_bottomStyle_2; 
                }
                else
                {
                    cell.CellStyle = sheet_2_bottomStyle_1; 
                }
                switch (run)
                {
                    case 0:
                        cell.SetCellValue("P'KGS");
                        break;

                    case 2:
                        cell.SetCellValue("SUB TOTAL:");
                        break;

                    case 3:
                    case 4:
                        break;

                    default:
                        {
                            String colum = Sheet_2_GetStringOfColum(run + 1);
                            String fomula = "SUM(" + colum + start_row + ":" + colum + end_row + ")";
                            cell.SetCellFormula(fomula);
                            break;
                        }
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

            int value = Sheet_2_size_2.Length > Sheet_2_size_1.Count ? Sheet_2_size_2.Length : Sheet_2_size_1.Count;

            value += 5/*before*/ + 4 /*after*/;

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

        private String Sheet_2_GetStringOfColum( int size)
        {
            byte[] data;

            int nguyen = size / 27;
            int le = 0;

            if (nguyen > 0)
            {
                le = size % 26;

                data = new byte[2];
                data[0] = (byte)(0x40 + (byte)nguyen);
                data[1] = (byte)(0x40 + (byte)le);
            }
            else
            {
                le = size % 27;
                data = new byte[1];
                data[0] = (byte)(0x40 + (byte)le);
            }

            return Encoding.ASCII.GetString(data); ;
        }
       
    }
}
