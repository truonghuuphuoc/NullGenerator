using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NULL_is_my_son
{
    class CombineEle
    {
        String size;
        double quality;

        public CombineEle(String size, double quality)
        {
            this.size = size;
            this.quality = quality;
        }

        public void SetSize(string size)
        {
            this.size = size;
        }

        public string GetSize()
        {
            return this.size;
        }

        public void SetQuality(double quality)
        {
            this.quality = quality;
        }

        public double GetQuality()
        {
            return this.quality;
        }
    }
    class CombineItem
    {
        String color_1;
        String color_2;

        List<CombineEle> listElement;

        public CombineItem(String color_1, String Color_2)
        {
            this.color_1 = color_1;
            this.color_2 = Color_2;

            listElement = new List<CombineEle>();
        }

        public String GetColor1()
        {
            return this.color_1;
        }

        public String GetColor2()
        {
            return this.color_2;
        }

        public List<CombineEle> GetListElement()
        {
            return this.listElement;
        }

        public void UpdateCombineElement(CombineEle element)
        {
            listElement.Add(element);
        }
    }
    class CombineMergerListItem
    {
        private List<CombineItem> listCombineItem;
        private PackingListItem packItem;

        public CombineMergerListItem(PackingListItem packItem)
        {
            this.packItem = packItem;
            listCombineItem = new List<CombineItem>();

        }

        public List<CombineItem> GetListCombineItem()
        {
            return listCombineItem;
        }

        public void RunCombineMergerListItem()
        {
            CombineItem main = new CombineItem(packItem.GetColor1(), packItem.GetColor2());
            main.UpdateCombineElement(new CombineEle(packItem.GetProductSize(), packItem.GetQuantity()));

            listCombineItem.Add(main);

            if(packItem.GetMergerList() != null && packItem.GetMergerList().Count > 0)
            {
                foreach(MergerItem merger in packItem.GetMergerList())
                {
                    bool iscombine = false;
                    foreach(CombineItem combine in  listCombineItem)
                    {
                        if(combine.GetColor1() == merger.GetColor1())
                        {
                            iscombine = true;
                            combine.UpdateCombineElement(new CombineEle(merger.GetProductSize(), merger.GetQuantity()));
                            break;
                        }
                        
                    }

                    if (!iscombine)
                    {
                        main = new CombineItem(merger.GetColor1(), merger.GetColor2());
                        main.UpdateCombineElement(new CombineEle(merger.GetProductSize(), merger.GetQuantity()));

                        listCombineItem.Add(main);
                    }
                }
            }
        }
    }
}
