using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NULL_is_my_son
{
    class MergerItem
    {
        public String nameProduct;
        public String color1;
        public String color2;
        public String productSize;
        public double quantity;

        public MergerItem()
        {
            this.nameProduct = "";
            this.color1 = "";
            this.color2 = "";
            this.productSize = "";
            this.quantity = 0;
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
    }
}
