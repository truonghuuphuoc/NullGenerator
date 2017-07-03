using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO; // File.Exists()

using NPOI.XSSF.UserModel; // XSSFWorkbook, XSSFSheet


namespace NULL_is_my_son
{
    class ParkingListParser
    {
        private List<ParkingParserItem> listPackingParser;
        private String dateTime;
        private String errorCode;

        private Database databaseInfor;

        public ParkingListParser(Database databaseInfor)
        {
            listPackingParser = null;
            dateTime = "";
            errorCode = "";

            this.databaseInfor = databaseInfor;
        }

        public String GetDateTime()
        {
            return dateTime;
        }

        public void SetDateTime(String value)
        {
            this.dateTime = value;
        }

        public List<ParkingParserItem> GetParkingParserList()
        {
            return this.listPackingParser;
        }

        public bool PackingListParser(String file)
        {
            bool addItem;
            PackingListItem item = null;

            if (!File.Exists(file))
            {
                this.errorCode = "Find not found\r\n";
                return false;
            }

            listPackingParser = new List<ParkingParserItem>();

            try
            {

                // get sheets list from xlsx
                FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read);

                XSSFWorkbook wb = new XSSFWorkbook(fs);

                if (wb.Count <= 0)
                {
                    this.errorCode = "Sheet in file is empty \r\n";
                    return false;
                }

                XSSFSheet sh = (XSSFSheet)wb.GetSheetAt(0);

                int index = 0;
                while (sh.GetRow(index) != null)
                {
                    //ignore
                    if (index == 0 || index == 1 || index == 2 )
                    {
                        index++;
                        continue;
                    }

                    //handle date time
                    if (index == 3)
                    {
                        if (sh.GetRow(index).Cells.Count >= 4)
                        {
                            var cell = sh.GetRow(index).GetCell(3);

                            if (cell != null && 
                                cell.CellType == NPOI.SS.UserModel.CellType.String &&
                                cell.StringCellValue != null)
                            {
                                int position = cell.StringCellValue.IndexOf("日期");

                                if (position >= 0)
                                {
                                    position = cell.StringCellValue.IndexOf(":");

                                    if (position >= 0)
                                    {
                                        this.dateTime = cell.StringCellValue.Substring(position + 1);
                                    }
                                }
                            }
                        }

                        index++;
                        continue;
                    }

                    //handle date time
                    if (index == 4)
                    {
                        if (this.dateTime == "")
                        {
                            if (sh.GetRow(index).Cells.Count >= 0)
                            {
                                var cell = sh.GetRow(index).GetCell(0);

                                if (cell != null &&
                                    cell.CellType == NPOI.SS.UserModel.CellType.String &&
                                    cell.StringCellValue != null)
                                {
                                    int position = cell.StringCellValue.IndexOf("主  旨");

                                    if (position >= 0)
                                    {
                                        position = cell.StringCellValue.IndexOf(":");
                                        int last_post = cell.StringCellValue.IndexOf("日");

                                        if (position >= 0 && (last_post-1) > position)
                                        {
                                            this.dateTime = cell.StringCellValue.Substring(position + 1, last_post - position);
                                        }
                                    }
                                }
                            }
                            
                        }

                        index++;
                        continue;
                    }

                    //ignore
                    if (index == 5)
                    {
                        index++;
                        continue;
                    }

                    addItem = false;


                    if (sh.GetRow(index).Cells.Count >= 5)
                    {
                        var cell = sh.GetRow(index).GetCell(4);

                        if (cell != null )
                        {
                            if (cell.CellType == NPOI.SS.UserModel.CellType.String &&
                                cell.StringCellValue != null)
                            {
                                addItem = true;

                                item = new PackingListItem();

                                item.SetNameProduct(cell.StringCellValue);
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                addItem = true;

                                item = new PackingListItem();

                                item.SetNameProduct(cell.NumericCellValue + "");
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Blank)
                            {
                                fs.Close();
                                return true;
                            }
                        }
                    }

                    if (!addItem)
                    {
                        index++;
                        continue;
                    }

                    if (sh.GetRow(index).Cells.Count >= 6)
                    {
                        var cell = sh.GetRow(index).GetCell(5);

                        if (cell != null)
                        {
                            if (cell.CellType == NPOI.SS.UserModel.CellType.String &&
                                cell.StringCellValue != null)
                            {
                                item.SetColor1(cell.StringCellValue);
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                item.SetColor1(cell.NumericCellValue + "");
                            }
                        }
                    }

                    if (sh.GetRow(index).Cells.Count >= 7)
                    {
                        var cell = sh.GetRow(index).GetCell(6);

                        if (cell != null)
                        {
                            if (cell.CellType == NPOI.SS.UserModel.CellType.String &&
                                cell.StringCellValue != null)
                            {
                                item.SetColor2(cell.StringCellValue);
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                item.SetColor2(cell.NumericCellValue + "");
                            }
                        }
                    }

                    if (sh.GetRow(index).Cells.Count >= 8)
                    {
                        var cell = sh.GetRow(index).GetCell(7);

                        if (cell != null)
                        {
                            if (cell.CellType == NPOI.SS.UserModel.CellType.String &&
                                cell.StringCellValue != null)
                            {
                                item.SetProductSize(cell.StringCellValue);
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                item.SetProductSize(cell.NumericCellValue + "");
                            }
                        }
                    }

                    if (sh.GetRow(index).Cells.Count >= 9)
                    {
                        var cell = sh.GetRow(index).GetCell(8);

                        if (cell != null)
                        {
                            if (cell.CellType == NPOI.SS.UserModel.CellType.String &&
                                cell.StringCellValue != null)
                            {
                                item.SetQuantity(Convert.ToDouble(cell.StringCellValue));
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                item.SetQuantity(cell.NumericCellValue);
                            }
                        }
                    }  

                    //add to list
                    AddPackingItem(item);
                    
                    index++;
                }


                fs.Close();
            }
            catch (IOException ex)
            {
                this.errorCode = ex.Message + "\r\n";

                return false;
            }

            return true;
        }

        public String GetErrorCode()
        {
            return this.errorCode;
        }

        public void MergerPackingList()
        {
            if (listPackingParser == null)
            {
                return;
            }

            foreach (ParkingParserItem parserItem in listPackingParser)
            {

                if (parserItem.GetListItem() == null)
                {
                    continue;
                }

                SliptToMerger(parserItem);
            }
        }

        ////////////////////////////////////////////////
        private void AddPackingItem(PackingListItem item)
        {
            ParkingParserItem parserItem = null;

            DatabaseItem dataItem = databaseInfor.FindItem(item.GetNameProduct());

            if (dataItem == null)
            {
                this.errorCode += item.GetNameProduct() + ": is empty in database\r\n";

                dataItem = new DatabaseItem();
            }

            if (listPackingParser == null)
            {
                listPackingParser = new List<ParkingParserItem>();
            }

            

            //add new
            if (listPackingParser.Count == 0)
            {
                parserItem = new ParkingParserItem(item.GetNameProduct(), item, dataItem);

                listPackingParser.Add(parserItem);

                return;
            }


            //add to exist
            for (int index = 0; index < listPackingParser.Count; index++)
            {
                if (item.GetNameProduct() == listPackingParser.ElementAt(index).GetNameProduct())
                {
                    listPackingParser.ElementAt(index).AddNewItemInList(item);

                    return;
                }
            }

            parserItem = new ParkingParserItem(item.GetNameProduct(), item, dataItem);
            //add new
            listPackingParser.Add(parserItem);

            return;
        }

        private double GetNumberOfPacket(double quantity, double maxsize)
        {
            int numberPakcet = (int)(quantity / maxsize);
            double phan_du = quantity - maxsize * numberPakcet;

            if (phan_du > 0)
            {
                numberPakcet += 1;
            }

            return numberPakcet;
        }

        private void SliptToMerger( ParkingParserItem parserItem)
        {
            PackingListItem packingItem;
            int base_index = 0;

            double totalQunatity = 0;
            double numberItem = 0;

            foreach (PackingListItem item in parserItem.GetListItem())
            {
                if (item.IsNeedMerger())
                {
                    totalQunatity += item.GetQuantity();
                    numberItem++;
                }
            }

            if (totalQunatity >= parserItem.GetDatabaseItem().GetMaxPacketSize())
            {

                double sliptQuantity = 0;
                PackingListItem itemAdd = null;

                for (int run = 0; run < parserItem.GetListItem().Count; run++)
                {
                    if (parserItem.GetListItem().ElementAt(run).IsNeedMerger())
                    {
                        base_index = run;
                        break;
                    }
                }

                //Need to slipt
                parserItem.GetListItem().ElementAt(base_index).SetNeedMerger(false);

                sliptQuantity = parserItem.GetDatabaseItem().GetMaxPacketSize() -
                                parserItem.GetListItem().ElementAt(base_index).GetQuantity();

                while (sliptQuantity > 0)
                {
                    bool found = false;

                    for (int run = base_index + 1; run < parserItem.GetListItem().Count; run++)
                    {
                        if (parserItem.GetListItem().ElementAt(run).IsNeedMerger())
                        {
                            if (parserItem.GetListItem().ElementAt(run).GetQuantity() > sliptQuantity)
                            {
                                found = true;

                                itemAdd = new PackingListItem();

                                itemAdd.SetColor1(parserItem.GetListItem().ElementAt(run).GetColor1());
                                itemAdd.SetColor2(parserItem.GetListItem().ElementAt(run).GetColor2());
                                itemAdd.SetNameProduct(parserItem.GetListItem().ElementAt(run).GetNameProduct());
                                itemAdd.SetNeedMerger(false);
                                itemAdd.SetProductSize(parserItem.GetListItem().ElementAt(run).GetProductSize());
                                itemAdd.SetQuantity(sliptQuantity);

                                parserItem.GetListItem().ElementAt(run).SetQuantity(parserItem.GetListItem().ElementAt(run).GetQuantity() - sliptQuantity);

                                sliptQuantity = 0;
                                parserItem.GetListItem().ElementAt(base_index).AddMergerItem(itemAdd);
                            }
                            else
                            {
                                found = true;

                                packingItem = parserItem.GetListItem().ElementAt(run);

                                packingItem.SetNeedMerger(false);

                                parserItem.GetListItem().RemoveAt(run);

                                parserItem.GetListItem().ElementAt(base_index).AddMergerItem(packingItem);

                                sliptQuantity -= packingItem.GetQuantity();
                            }

                            break;
                        }
                    }

                    if (!found)
                    {
                        break;
                    }
                }

                SliptToMerger(parserItem);
            }
            else /*less than*/ if (numberItem != 0)
            {
                for (int run = 0; run < parserItem.GetListItem().Count; run++)
                {
                    if (parserItem.GetListItem().ElementAt(run).IsNeedMerger())
                    {
                        parserItem.GetListItem().ElementAt(run).SetNeedMerger(false);
                        base_index = run;
                        break;
                    }
                }

                numberItem -= 1;

                while (numberItem > 0)
                {
                    packingItem = null;

                    if (base_index >= parserItem.GetListItem().Count)
                    {
                        break;
                    }

                    for (int run = base_index; run < parserItem.GetListItem().Count; run++)
                    {
                        if (parserItem.GetListItem().ElementAt(run).IsNeedMerger())
                        {
                            packingItem = parserItem.GetListItem().ElementAt(run);

                            packingItem.SetNeedMerger(false);

                            parserItem.GetListItem().RemoveAt(run);
                            break;
                        }
                    }

                    if (packingItem != null)
                    {
                        parserItem.GetListItem().ElementAt(base_index).AddMergerItem(packingItem);
                    }

                    numberItem -= 1;
                }
            }
        }
    }
}
