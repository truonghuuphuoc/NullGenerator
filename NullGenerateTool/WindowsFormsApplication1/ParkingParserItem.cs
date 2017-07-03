using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NULL_is_my_son
{
    class ParkingParserItem
    {
        private String nameProduct;
        private List<PackingListItem> listItem;
        private DatabaseItem databaseItem;

        public ParkingParserItem(String name, PackingListItem item, DatabaseItem databaseItem)
        {
            this.nameProduct = name;
            this.databaseItem = databaseItem;

            this.listItem = new List<PackingListItem>();
            if (item != null)
            {
                int phan_nguyen = (int)(item.GetQuantity() / databaseItem.GetMaxPacketSize());
                int phan_du = (int)(item.GetQuantity()) - (int)(phan_nguyen * databaseItem.GetMaxPacketSize());

                if (phan_du == 0 || phan_nguyen == 0)
                {
                    if (phan_nguyen == 0)
                    {
                        item.SetNeedMerger(true);
                    }
                    else
                    {
                        item.SetNeedMerger(false);
                    }
                    listItem.Add(item);
                }
                else
                {
                    //add phan nguyen
                    item.SetQuantity(phan_nguyen * databaseItem.GetMaxPacketSize());
                    item.SetNeedMerger(false);
                    listItem.Add(item);

                    //add phan du
                    PackingListItem item_du = new PackingListItem();

                    item_du.SetColor1(item.GetColor1());
                    item_du.SetColor2(item.GetColor2());
                    item_du.SetNameProduct(item.GetNameProduct());
                    item_du.SetProductSize(item.GetProductSize());
                    item_du.SetQuantity(phan_du);

                    item_du.SetNeedMerger(true);
                    listItem.Add(item_du);
                }
            }
        }

        public String GetNameProduct()
        {
            return this.nameProduct;
        }

        public void SetNameProduct(String value)
        {
            this.nameProduct = value;
        }

        public DatabaseItem GetDatabaseItem()
        {
            return this.databaseItem;
        }

        public void SetDatabaseItem(DatabaseItem value)
        {
            this.databaseItem = value;
        }

        public void AddNewItemInList(PackingListItem item)
        {
            if (listItem == null)
            {
                listItem = new List<PackingListItem>();
            }

            if (item != null && this.databaseItem != null)
            {
                int phan_nguyen = (int)(item.GetQuantity() / databaseItem.GetMaxPacketSize());
                int phan_du = (int)(item.GetQuantity()) - (int)(phan_nguyen * databaseItem.GetMaxPacketSize());

                if (phan_du == 0 || phan_nguyen == 0)
                {
                    if (phan_nguyen == 0)
                    {
                        item.SetNeedMerger(true);
                    }
                    else
                    {
                        item.SetNeedMerger(false);
                    }

                    listItem.Add(item);
                }
                else
                {
                    //add phan nguyen
                    item.SetQuantity(phan_nguyen * databaseItem.GetMaxPacketSize());
                    item.SetNeedMerger(false);
                    listItem.Add(item);

                    //add phan du
                    PackingListItem item_du = new PackingListItem();

                    item_du.SetColor1(item.GetColor1());
                    item_du.SetColor2(item.GetColor2());
                    item_du.SetNameProduct(item.GetNameProduct());
                    item_du.SetProductSize(item.GetProductSize());
                    item_du.SetQuantity(phan_du);

                    item_du.SetNeedMerger(true);
                    listItem.Add(item_du);
                }
            }
        }

        public List<PackingListItem> GetListItem()
        {
            return this.listItem;
        }
    }
}
