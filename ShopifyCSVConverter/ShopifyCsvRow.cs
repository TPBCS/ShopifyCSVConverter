using System;
using System.Linq;
using CsvHelper.Configuration;
using System.Data;

namespace ShopifyCSVConverter
{
    public class ShopifyCsvRow : IDisposable
    {   
        public object Handle => data[0];
        public object Title => data[1];
        public object BodyHtml => data[2];
        public object Vendor => data[3];
        public object Type => data[4];
        public object Tags => data[5];
        public object Published => data[6];
        public object Option1Name => data[7];
        public object Option1Value => data[8];
        public object Option2Name => data[9];
        public object Option2Value => data[10];
        public object Option3Name => data[11];
        public object Option3Value => data[12];
        public object SKU => data[13];
        public object VariantGrams => data[14];
        public object VariantInventoryTracker => data[15];
        public object VariantInventoryQty => data[16];
        public object VariantInventoryPolicy => data[17];
        public object VariantFulfilmentService => data[18];
        public object VariantPrice => data[19];
        public object VariantCompareAtPrice => data[20];
        public object VarianRequiresShipping => data[21];
        public object VariantTaxable => data[22];
        public object VariantBarcode => data[23];
        public object ImageSrc => data[24];
        public object ImageAltText => data[25];
        public object GiftCard => data[26];
        public object GoogleShoppingMPN => data[27];
        public object GoogleShoppingAgeGroup => data[28];
        public object GoogleShoppingGender => data[29];
        public object GoogleShoppingGoogleProductCategory => data[30];
        public object SEOTitle => data[31];
        public object SEODescription => data[32];
        public object GoogleShoppingAdWordsGrouping => data[33];
        public object GoogleShoppingAdWordsLabels => data[34];
        public object GoogleShoppingCondition => data[35];
        public object GoogleShoppingCustomProduct => data[36];
        public object GoogleShoppingCustomLabel0 => data[37];
        public object GoogleShoppingCustomLabel1 => data[38];
        public object GoogleShoppingCustomLabel2 => data[39];
        public object GoogleShoppingCustomLabel3 => data[40];
        public object GoogleShoppingCustomLabel4 => data[41];
        public object VariantImage => data[42];
        public object VariantWeightUnit => data[43];
        public object Collection => data[44];
        private object[] data;

        public ShopifyCsvRow(DataRow row)
        {
            data = new object[45];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = row[i];
            }
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing) data.ToList().Clear();
                disposedValue = true;
            }
        }
        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }

    public class ShopifyCsvRowMap : ClassMap<ShopifyCsvRow>
    {
        public ShopifyCsvRowMap()
        {
            this.Map(m => m.Handle ).Name("Handle").Index(0);
            this.Map(m => m.Title ).Name("Title").Index(1);
            this.Map(m => m.BodyHtml ).Name("Body (HTML)").Index(2);
            this.Map(m => m.Vendor ).Name("Vendor").Index(3);
            this.Map(m => m.Type ).Name("Type").Index(4);
            this.Map(m => m.Tags ).Name("Tags").Index(5);
            this.Map(m => m.Published ).Name("Published").Index(6);
            this.Map(m => m.Option1Name ).Name("Option1 Name").Index(7);
            this.Map(m => m.Option1Value ).Name("Option1 Value").Index(8);
            this.Map(m => m.Option2Name ).Name("Option2 Name").Index(9);
            this.Map(m => m.Option2Value ).Name("Option2 Value").Index(10);
            this.Map(m => m.Option3Name ).Name("Option3 Name").Index(11);
            this.Map(m => m.Option3Value ).Name("Option3 Value").Index(12);
            this.Map(m => m.SKU ).Name("SKU").Index(13);
            this.Map(m => m.VariantGrams ).Name("Variant Grams").Index(14);
            this.Map(m => m.VariantInventoryTracker ).Name("Variant Inventory Tracker").Index(15);
            this.Map(m => m.VariantInventoryQty ).Name("Variant Inventory Quantity").Index(16);
            this.Map(m => m.VariantInventoryPolicy ).Name("Variant Inventory Policy").Index(17);
            this.Map(m => m.VariantFulfilmentService ).Name("Variant Fulfillment Service").Index(18);
            this.Map(m => m.VariantPrice ).Name("Variant Price").Index(19);
            this.Map(m => m.VariantCompareAtPrice ).Name("Variant Compare At Price").Index(20);
            this.Map(m => m.VarianRequiresShipping ).Name("Variant Requires Shipping").Index(21);
            this.Map(m => m.VariantTaxable ).Name("Variant Taxable").Index(22);
            this.Map(m => m.VariantBarcode ).Name("Variant Barcode").Index(23);
            this.Map(m => m.ImageSrc ).Name("Image Src").Index(24);
            this.Map(m => m.ImageAltText ).Name("Image Alt Text").Index(25);
            this.Map(m => m.GiftCard ).Name("Gift Card").Index(26);
            this.Map(m => m.GoogleShoppingMPN ).Name("Google Shopping / MPN").Index(27);
            this.Map(m => m.GoogleShoppingAgeGroup ).Name("Google Shopping / Age Group").Index(28);
            this.Map(m => m.GoogleShoppingGender ).Name("Google Shopping / Gender").Index(29);
            this.Map(m => m.GoogleShoppingGoogleProductCategory ).Name("Google Shopping / Google Product Category").Index(30);
            this.Map(m => m.SEOTitle ).Name("SEO Title").Index(31);
            this.Map(m => m.SEODescription ).Name("SEO Description").Index(32);
            this.Map(m => m.GoogleShoppingAdWordsGrouping ).Name("Google Shopping / AdWords Grouping").Index(33);
            this.Map(m => m.GoogleShoppingAdWordsLabels ).Name("Google Shopping / AdWords Labels").Index(34);
            this.Map(m => m.GoogleShoppingCondition ).Name("Google Shopping / Condition").Index(35);
            this.Map(m => m.GoogleShoppingCustomProduct ).Name("Google Shopping / Custom Product").Index(36);
            this.Map(m => m.GoogleShoppingCustomLabel0 ).Name("Google Shopping / Custom Label 0").Index(37);
            this.Map(m => m.GoogleShoppingCustomLabel1 ).Name("Google Shopping / Custom Label 1").Index(38);
            this.Map(m => m.GoogleShoppingCustomLabel2 ).Name("Google Shopping / Custom Label 2").Index(39);
            this.Map(m => m.GoogleShoppingCustomLabel3 ).Name("Google Shopping / Custom Label 3").Index(40);
            this.Map(m => m.GoogleShoppingCustomLabel4 ).Name("Google Shopping / Custom Label 4").Index(41);
            this.Map(m => m.VariantImage ).Name("Variant Image").Index(42);
            this.Map(m => m.VariantWeightUnit ).Name("Variant Weight Unit").Index(43);
            this.Map(m => m.Collection ).Name("Collection").Index(44);
        }
    }
}
