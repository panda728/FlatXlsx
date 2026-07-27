namespace BenchmarkSample
{
    /// <summary>The fixed-width sample file both benchmarks are built from.</summary>
    static class SampleData
    {
        const int HEADER_LEN = 9 + 30;
        const int DETAIL_LEN = 10;
        const int DETAIL_COUNT = 10;

        /// <summary>Column order follows Row's property declaration order so that every
        /// benchmark produces the same layout.</summary>
        public static readonly string[] ColumnTitles =
        {
            "行番号", "ヘッダID", "明細ID", "ヘッダ1", "ヘッダ2", "ヘッダ3", "ヘッダ4", "ヘッダ5",
            "ヘッダ6", "ヘッダ7", "明細データ", "フッタ1", "フッタ2", "フッタ3", "フッタ4",
            "フッタ5", "フッタ6", "フッタ7", "フッタ8",
        };

        /// <summary>Reads <c>data01.dat</c> into rows. Every line of the file expands into
        /// <see cref="DETAIL_COUNT"/> detail rows, so the file's 10 lines make one 100-row block.</summary>
        public static List<Row> LoadBlock()
        {
            var list = new List<Row>();
            var lineNum = 0;
            using var sr = new StreamReader("data01.dat");
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                if (line == null)
                    break;

                if (!int.TryParse(line[..9], out var headerID))
                    throw new ApplicationException("Could not be converted to int.");
                for (int i = 0; i < DETAIL_COUNT; i++)
                {
                    list.Add(new Row
                    {
                        LineNum = lineNum++,
                        HeaderID = headerID,
                        DetailID = i + 1,
                        Data = line.Substring(HEADER_LEN + (DETAIL_LEN * i), DETAIL_LEN),
                        Header01 = line[9..13],
                        Header02 = line[13..20],
                        Header03 = line[20..25],
                        Header04 = line[25..27],
                        Header05 = line[27..30],
                        Header06 = line[30..33],
                        Header07 = line[33..39],
                        Footer01 = line[139..144],
                        Footer02 = line[144..153],
                        Footer03 = line[153..155],
                        Footer04 = line[155..158],
                        Footer05 = line[158..166],
                        Footer06 = line[166..171],
                        Footer07 = line[171..174],
                        Footer08 = line[174..178],
                    });
                }
            }
            return list;
        }
    }
}
