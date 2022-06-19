using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
    internal class ResultTableRow
    {
        /// <summary>
        /// Артикул 1
        /// </summary>
        public string VendorCode1 { get; set; } = string.Empty;
        /// <summary>
        /// Артикул 2
        /// </summary>
        public string VendorCode2 { get; set; } = string.Empty;
        /// <summary>
        /// Наименование
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// Количество
        /// </summary>
        public int Count { get; set; } = 0;
        /// <summary>
        /// Штрих-код
        /// </summary>
        public string Barcode { get; set; } = "000000000000";
    }
}
