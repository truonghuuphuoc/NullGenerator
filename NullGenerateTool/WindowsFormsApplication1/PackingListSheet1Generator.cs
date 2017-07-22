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
    class PackingListSheet1Generator
    {
        private String[] Sheet_1_Header = {   "日期", "貨號",     "顏色", "顏色2", "尺寸", 
                                                                         "QTY",  "箱數",     "箱號", "合計",  "淨重", 
                                                                         "毛重", "外箱規格", "備註"};

        private XSSFWorkbook wb;
        private XSSFSheet sheet_1;
        private int rowIndex_sheet_1;

        private double numberPacket_sheet_1;

        private ICellStyle sheet_1_headerStyle_2;
        private ICellStyle sheet_1_headerStyle;


        private Database databaseInfor;
        private ParkingListParser packingParser;

        public PackingListSheet1Generator(XSSFWorkbook wb, XSSFSheet sheet_1, Database databaseInfor, ParkingListParser packingParser)
        {
            this.wb = wb;
            this.sheet_1 = sheet_1;
            this.databaseInfor = databaseInfor;
            this.packingParser = packingParser;

            rowIndex_sheet_1 = 0;

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


            var r = sheet_1.CreateRow(rowIndex_sheet_1);
            rowIndex_sheet_1++;

            for (int idex = 0; idex < Sheet_1_Header.Length; idex++)
            {
                var cell = r.CreateCell(idex);
                cell.SetCellValue(Sheet_1_Header[idex]);
                cell.CellStyle = sheet_1_headerStyle;
            }
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
                            if (valueMerger > item.GetQuantity())
                            {
                                cell.CellStyle = sheet_1_headerStyle_2;
                            }
                            break;

                        case 6:
                            cell.SetCellValue(1);
                            break;

                        case 7:
                            cell.SetCellValue(numberPacket_sheet_1);
                            break;

                        case 8:
                            if (valueMerger <= item.GetQuantity())
                            {
                                cell.SetCellValue(valueMerger);
                            }
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
                for (int tmp = 0; tmp < (int)(item.GetQuantity() / databaseItem.GetMaxPacketSize()); tmp++)
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
                                //cell.SetCellValue(databaseItem.GetNetWeight());
                                break;

                            case 10:
                                //cell.SetCellValue(databaseItem.GetAllWeight());
                                break;

                            case 11:
                                //cell.SetCellValue(databaseItem.GetPacketInformation());
                                break;

                            default:
                                break;
                        }
                    }
                }

                r= sheet_1.GetRow(rowIndex_sheet_1 - (1 + item.GetMergerList().Count));
                String fomular = "SUM(F" + (rowIndex_sheet_1 -  item.GetMergerList().Count) + ":F" + rowIndex_sheet_1 + ")";
                if(r.GetCell(8) != null)
                {
                    r.GetCell(8).CellFormula = fomular;
                }

                NPOI.SS.Util.CellRangeAddress range_1 = new NPOI.SS.Util.CellRangeAddress(rowIndex_sheet_1 - (1 + item.GetMergerList().Count),
                                                                                          rowIndex_sheet_1 - 1, 6, 6);

                NPOI.SS.Util.CellRangeAddress range_2 = new NPOI.SS.Util.CellRangeAddress(rowIndex_sheet_1 - (1 + item.GetMergerList().Count),
                                                                                          rowIndex_sheet_1 - 1, 8, 8);
                sheet_1.AddMergedRegion(range_1);
                sheet_1.AddMergedRegion(range_2);

            }
        }


        public void GeneratePackingListSheet_1()
        {
             if (packingParser.GetParkingParserList() != null
                    && packingParser.GetParkingParserList().Count > 0)
                {
                    numberPacket_sheet_1 = 0;

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
                }
        }

    }
}
