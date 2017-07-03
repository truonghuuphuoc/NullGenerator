using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NULL_is_my_son
{
    class DatabaseItem
    {
        public String nameProduct;

        public String GetNameProduct()
        {
            return nameProduct;
        }

        public void SetNameProduct(String name)
        {
            this.nameProduct = name;
        }

       

        public double maxPacketSize;

        public double GetMaxPacketSize()
        {
            return this.maxPacketSize;
        }
        public void SetMaxPacketSize(double value)
        {
            this.maxPacketSize = value;
        }

        public double netWeight;

        public double GetNetWeight()
        {
            return this.netWeight;
        }
        public void SetNetWeight(double value)
        {
            this.netWeight = value;
        }

        public double allWeight;

        public double GetAllWeight()
        {
            return this.allWeight;
        }
        public void SetAllWeight(double value)
        {
            this.allWeight = value;
        }
        public String packetInformation;

        public String GetPacketInformation()
        {
            return this.packetInformation;
        }
        public void SetPacketInformation(String value)
        {
            this.packetInformation = value;
        }

        public DatabaseItem()
        {
            nameProduct = "";

            maxPacketSize = 1;
            netWeight = 0;
            allWeight = 0;

            packetInformation = "";
        }

        public void SetDataValue(int index, String value_str)
        {
            switch (index)
            {
                case 0:
                    this.nameProduct = value_str;
                    break;
                    
                case 1:
                    this.maxPacketSize = Convert.ToDouble(value_str);
                    break;

                case 2:
                    this.netWeight = Convert.ToDouble(value_str);
                    break;

                case 3:
                    this.allWeight = Convert.ToDouble(value_str);
                    break;

                case 4:
                    this.packetInformation = value_str;
                    break;

                default:
                    break;
            }
        }
    }
}
