using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NULL_is_my_son
{
    class PackingListItem
    {
        public String nameProduct;
        public String color1;
        public String color2;
        public String productSize;
        public double quantity;

        public bool isNeedMerger;

        public List<MergerItem> mergerList;

        public PackingListItem()
        {
            this.nameProduct = "";
            this.color1 = "";
            this.color2 = "";
            this.productSize = "";
            this.quantity = 0;

            this.isNeedMerger = false;

            this.mergerList = null;
        }


        public String GetNameProduct()
        {
            return this.nameProduct;
        }

        public void SetNameProduct(String value)
        {
            this.nameProduct = value;
        }

        public String GetColor1()
        {
            return this.color1;
        }

        public void SetColor1(String value)
        {
            this.color1 = value;
        }

        public String GetColor2()
        {
            return this.color2;
        }

        public void SetColor2(String value)
        {
            this.color2 = value;
        }

        public String GetProductSize()
        {
            return this.productSize;
        }

        public void SetProductSize(String value)
        {
            this.productSize = value;
        }

        public double GetQuantity()
        {
            return this.quantity;
        }

        public void SetQuantity(double value)
        {
            this.quantity = value;
        }

        public bool IsNeedMerger()
        {
            return this.isNeedMerger;
        }

        public void SetNeedMerger(bool value)
        {
            this.isNeedMerger = value;
        }

        public List<MergerItem> GetMergerList()
        {
            return this.mergerList;
        }

        public void AddMergerItem(PackingListItem packingItem)
        {
            MergerItem item = new MergerItem();

            item.SetColor1(packingItem.GetColor1());
            item.SetColor2(packingItem.GetColor2());

            item.SetNameProduct(packingItem.GetNameProduct());
            item.SetProductSize(packingItem.GetProductSize());

            item.SetQuantity(packingItem.GetQuantity());

            if (this.mergerList == null)
            {
                this.mergerList = new List<MergerItem>();
            }

            this.mergerList.Add(item);
        }

    }
}
