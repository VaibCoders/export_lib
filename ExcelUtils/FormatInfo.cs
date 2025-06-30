using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUtils
{
    /// <summary>Описание формата диапазона ячеек</summary>
    public class FormatInfo
    {
        /// <summary>Выравнивание по горизонтали</summary>
        public ColumnAlignment Alignment { get; set; } = ColumnAlignment.General;

        /// <summary>Цвет фона</summary>
        public int FillColor { get; set; }

        /// <summary>Жирность текста шрифта</summary>
        public bool FontBold { get; set; } = false;
        /// <summary>Цвет текста шрифта</summary>
        public int FontColor { get; set; } 
        /// <summary>Наклон текста шрифта</summary>
        public bool FontItalic { get; set; } = false;
        /// <summary>Имя текста шрифта</summary>
        public string FontName { get; set; } = "Calibri";
        /// <summary>Размер текста шрифта</summary>
        public float FontSize { get; set; } = 11;
        /// <summary>Зачеркивание текста шрифта</summary>
        public bool FontStrikethrough { get; set; } = false;
        /// <summary>Подчеркивание текста шрифта</summary>
        public bool FontUnderline { get; set; } = false;

        /// <summary>Формула диапазона ячеек в формате Excel</summary>
        public string? Formula { get; set; }
        /// <summary>Форма значения диапазона ячеек в формате Excel</summary>
        public string NumberFormat { get; set; } = "";
    }

    public enum ColumnAlignment
    {
        /// <summary>По умолчанию зависит от типа ячейки</summary>
        General = 0,
        /// <summary>Слева</summary>
        Left = 1,
        /// <summary>По центру</summary>
        Center = 2,
        /// <summary>Справа</summary>
        Right = 3
    }

}
