using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShopifyCSVConverter
{
    class ShopifyCsv
    {
        public Dictionary<int, Row> Rows;

        private static string[] Columns = new string[]
        {
            "Handle",
            "Title",
            "Body (HTML)",
            "Vendor",
            "Type",
            "Tags",
            "Published",
            "Option 1 Name",
            "Option 1 Value",
            "Option 2 Name",
            "Option 2 Value",
            "Option 3 Name",
            "Option 3 Value",
            "SKU",
            "Variant Grams",
            "Variant Inventory Tracker",
            "Variant Inventory Qty",
            "Variant Inventory Policy",
            "Variant Fulfilment Service",
            "Variant Price",
            "Variant Compare At Price",
            "Varian Requires Shipping",
            "Variant Taxable",
            "Variant Barcode",
            "Image Src",
            "Image Alt Text",
            "GiftCard",
            "Google Shopping / MPN",
            "Google Shopping / Age Group",
            "Google Shopping / Gender",
            "Google Shopping / Google Product Category",
            "SEO Title",
            "SEO Description",
            "Google Shopping / AdWords Grouping",
            "Google Shopping / AdWords Labels",
            "Google Shopping / Condition",
            "Google Shopping / Custom Product",
            "Google Shopping / Custom Label 0",
            "Google Shopping / Custom Label 1",
            "Google Shopping / Custom Label 2",
            "Google Shopping / Custom Label 3",
            "Google Shopping / Custom Label 4",
            "Variant Image",
            "Variant Weight Unit",
            "Collection"
        };

        public ShopifyCsv()
        {

        }

        public class Row
        {
            public string Handle => (string)rowData[Columns[0]];
            public string Title => (string)rowData[Columns[1]];
            public string BodyHtml => (string)rowData[Columns[2]];
            public string Vendor => (string)rowData[Columns[3]];
            public string Type => (string)rowData[Columns[4]];
            public string Tags => (string)rowData[Columns[5]];
            public string Published => (string)rowData[Columns[6]];
            public string Option1Name => (string)rowData[Columns[7]];
            public string Option1Value => (string)rowData[Columns[8]];
            public string Option2Name => (string)rowData[Columns[9]];
            public string Option2Value => (string)rowData[Columns[10]];
            public string Option3Name => (string)rowData[Columns[11]];
            public string Option3Value => (string)rowData[Columns[12]];
            public string SKU => (string)rowData[Columns[13]];
            public string VariantGrams => (string)rowData[Columns[14]];
            public string VariantInventoryTracker => (string)rowData[Columns[15]];
            public string VariantInventoryQty => (string)rowData[Columns[16]];
            public string VariantInventoryPolicy => (string)rowData[Columns[17]];
            public string VariantFulfilmentService => (string)rowData[Columns[18]];
            public string VariantPrice => (string)rowData[Columns[19]];
            public string VariantCompareAtPrice => (string)rowData[Columns[20]];
            public string VarianRequiresShipping => (string)rowData[Columns[21]];
            public string VariantTaxable => (string)rowData[Columns[22]];
            public string VariantBarcode => (string)rowData[Columns[23]];
            public string ImageSrc => (string)rowData[Columns[24]];
            public string ImageAltText => (string)rowData[Columns[25]];
            public string GiftCard => (string)rowData[Columns[26]];
            public string GoogleShoppingMPN => (string)rowData[Columns[27]];
            public string GoogleShoppingAgeGroup => (string)rowData[Columns[28]];
            public string GoogleShoppingGender => (string)rowData[Columns[29]];
            public string GoogleShoppingGoogleProductCategory => (string)rowData[Columns[30]];
            public string SEOTitle => (string)rowData[Columns[31]];
            public string SEODescription => (string)rowData[Columns[32]];
            public string GoogleShoppingAdWordsGrouping => (string)rowData[Columns[33]];
            public string GoogleShoppingAdWordsLabels => (string)rowData[Columns[34]];
            public string GoogleShoppingCondition => (string)rowData[Columns[35]];
            public string GoogleShoppingCustomProduct => (string)rowData[Columns[36]];
            public string GoogleShoppingCustomLabel0 => (string)rowData[Columns[37]];
            public string GoogleShoppingCustomLabel1 => (string)rowData[Columns[38]];
            public string GoogleShoppingCustomLabel2 => (string)rowData[Columns[39]];
            public string GoogleShoppingCustomLabel3 => (string)rowData[Columns[40]];
            public string GoogleShoppingCustomLabel4 => (string)rowData[Columns[41]];
            public string VariantImage => (string)rowData[Columns[42]];
            public string VariantWeightUnit => (string)rowData[Columns[43]];
            public string Collection => (string)rowData[Columns[44]];
            
            private Dictionary<string, string> rowData;

            public Row()
            {

            }

            public Row(string[] data)
            {
                rowData = new Dictionary<string, string>();
                for (int i = 0; i < data.Length; i++)
                {
                    rowData.Add(Columns[i], data[i] == null ? "" : data[i]);
                }
            }

            public Row(Dictionary<string, string> data)
            {
                rowData = data;
            }

            public string NameOf(int index)
            {
                return Columns[index];
            }

            public int IndexOf(string name)
            {
                return Columns.ToList().IndexOf(name);
            }
        }
    }

    
}
