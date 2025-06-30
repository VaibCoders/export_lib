using System.Data;

namespace ExcelUtils
{
    /// <summary>Описание таблицы на листе</summary>
    public class TableInfo
    {

        /// <summary>Конструктор</summary>
        public TableInfo()
        {
            Columns = new List<ColumnInfo>();
            Groups = new List<string>();
            ConditionFormattings = new List<ConditionFormattingTableInfo>();
        }

        /// <summary>Условное форматирование таблицы</summary>
        public List<ConditionFormattingTableInfo> ConditionFormattings { get; set; }

        /// <summary>Список описаний столбцов таблицы</summary>
        public List<ColumnInfo> Columns { get; set; }

        /// <summary>Список столбцов для группировки</summary>
        public List<string> Groups { get; set; }

        /// <summary>Список столбцов для сортировки. Названия столбцов через "," (после имени можно указать desc или asc)</summary>
        public string? ColumnSort { get; set; }

        /// <summary>Создание столбцов на основе данных таблицы</summary>
        public void ColumnsCreate(DataTable Table)
        {
            Columns.Clear();
            foreach (DataColumn col in Table.Columns)
                Columns.Add(new ColumnInfo(col.ColumnName));
        }


        /// <summary>Описание условного форматирования таблицы</summary>
        public class ConditionFormattingTableInfo
        {
            /// <summary>Конструктор</summary>
            /// <param name="ConditionOperator">Условие сравнения</param>
            /// <param name="ConditionValue1">Значение 1</param>
            /// <param name="ConditionValue2">Значение 2</param>
            public ConditionFormattingTableInfo(Conditions ConditionOperator, object? ConditionValue1 = null, object? ConditionValue2 = null)
            {
                this.ConditionOperator = ConditionOperator;
                if (ConditionValue1 != null) this.ConditionValue1 = ConditionValue1;
                if (ConditionValue2 != null) this.ConditionValue2 = ConditionValue2;
                SourceColumn = "";
            }

            /// <summary>Условие сравнения</summary>
            public Conditions ConditionOperator { get; set; }
            /// <summary>Значение 1</summary>
            public dynamic? ConditionValue1 { get; set; }
            /// <summary>Значение 2</summary>
            public dynamic? ConditionValue2 { get; set; }
            /// <summary>Формат</summary>
            public FormatInfo Formating { get; set; } = new FormatInfo();
            /// <summary>Столбец источник</summary>
            public string SourceColumn { get; set; }
            /// <summary>Столбец цель. Если не задан, то форматирование применяется ко всей строке</summary>
            public string? TargetColumn { get; set; }
        }

        public enum Conditions : int
        {
            // Expression
            Equal = 0,
            NotEqual = 1,
            Greater = 2,
            Less = 3,
            GreaterOrEqual = 4,
            LessOrEqual = 5,

            // Range
            Between = 6,
            NotBetween = 7,

            // Text
            Contains = 8,
            NotContains = 9,
            BeginsWith = 10,
            EndsWith = 11,

            // Special
            ContainNulls = 12,
            ContainNonNulls = 13,
            ContainBlanks = 14,
            ContainNonBlanks = 15
        }

    }

}
