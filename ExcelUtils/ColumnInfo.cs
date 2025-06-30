namespace ExcelUtils
{
    /// <summary>Описание столбцов таблицы</summary>
    public class ColumnInfo
    {
        /// <summary>Конструктор</summary>
        /// <param name="ColumnName">Уникальное имя столбца</param>
        public ColumnInfo(string ColumnName)
        {
            this.ColumnName = ColumnName;
            this.ColumnCaption = ColumnName;
            Formating = new FormatInfo();
            FormatingHeader = new FormatInfo();
            SummatyList = new List<SummaryTypes>();
        }

        ///// <summary>Условное форматирование диапазона ячеек</summary>
        //public List<ConditionFormattingColumnInfo> ConditionFormattings { get; set; } = new List<ConditionFormattingColumnInfo>();

        /// <summary>Имя столбца</summary>
        public string ColumnName { get; set; }

        /// <summary>Заголовок столбца</summary>
        public string ColumnCaption { get; set; }

        /// <summary>Список столбцов для сортировки (в виде "ColName1, ColName1 desc")</summary>
        public string? ColumnFilter { get; set; }

        /// <summary>Видимость столбца</summary>
        public bool Visible { get; set; } = true;

        /// <summary>Ширина столбца ()</summary>
        public double Width { get; set; }

        /// <summary>Формат ячеек столбца</summary>
        public FormatInfo Formating { get; set; }
        /// <summary>Формат заголовок столбца</summary>
        public FormatInfo FormatingHeader { get; set; }

        /// <summary>Итоги столбцов</summary>
        public List<SummaryTypes> SummatyList { get; set; }
    }

    /// <summary>Типы агрегатных функций</summary>
    public enum SummaryTypes: int
    {
        // The sum of all values in a column.
        Sum=0,
        // The minimum value in a column.
        Min=1,
        // The maximum value in a column.
        Max=2,
        // The record count.
        Count=3,
        // The average value of a column.
        Average=4,
        // Specifies whether calculations should be performed manually using a specially designed event.
        Custom=5,
        // Disables summary value calculation.
        None=6
    }

}
