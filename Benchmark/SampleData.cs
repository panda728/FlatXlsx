using System.Globalization;

namespace BenchmarkSample
{
    /// <summary>
    /// Builds the rows every benchmark measures, with the properties that drive an export's cost
    /// stated as arguments rather than baked into a data file.
    /// </summary>
    /// <remarks>
    /// What matters to a spreadsheet writer is not how realistic a value looks but how many rows
    /// there are, how wide the values are, and - the one that is easy to get wrong - how many of
    /// them are <em>distinct</em>, because distinct strings are what the shared-string table
    /// grows on. Those are parameters here, so a benchmark can say which one it is varying.
    ///
    /// Nothing is random: the same arguments produce the same rows on every machine and every
    /// run, so results stay comparable without having to pin a seed.
    /// </remarks>
    public static class SampleData
    {
        /// <summary>Column order follows Row's property declaration order so that every
        /// benchmark produces the same layout.</summary>
        public static readonly string[] ColumnTitles =
        {
            "行番号", "伝票番号", "明細番号",
            "拠点", "取引先", "商品", "単位", "状態", "経路", "担当",
            "明細キー",
            "区分1", "区分2", "区分3", "区分4", "区分5", "区分6", "区分7", "区分8",
        };

        /// <summary>A column whose values repeat across the export - a branch code, a status, a
        /// unit. <paramref name="count"/> is how many distinct values it takes.</summary>
        static string[] Pool(string prefix, int count, int width)
        {
            var pool = new string[count];
            for (var i = 0; i < count; i++)
            {
                var body = i.ToString(CultureInfo.InvariantCulture).PadLeft(width - prefix.Length, '0');
                pool[i] = prefix + body;
            }
            return pool;
        }

        //                                          distinct  width
        static readonly string[] _branch = Pool("BR", 5, 4);
        static readonly string[] _client = Pool("CL", 40, 7);
        static readonly string[] _product = Pool("PR", 120, 5);
        static readonly string[] _unit = Pool("U", 4, 2);
        static readonly string[] _state = Pool("S", 6, 3);
        static readonly string[] _route = Pool("R", 3, 3);
        static readonly string[] _staff = Pool("SF", 25, 6);
        static readonly string[] _kind1 = Pool("KA", 8, 5);
        static readonly string[] _kind2 = Pool("KB", 60, 9);
        static readonly string[] _kind3 = Pool("K", 4, 2);
        static readonly string[] _kind4 = Pool("K", 7, 3);
        static readonly string[] _kind5 = Pool("KE", 30, 8);
        static readonly string[] _kind6 = Pool("KF", 12, 5);
        static readonly string[] _kind7 = Pool("K", 5, 3);
        static readonly string[] _kind8 = Pool("KH", 9, 4);

        /// <summary>19 columns: 3 numbers, 15 repeating text columns, and one key column whose
        /// number of distinct values is <paramref name="distinctKeys"/>.</summary>
        /// <param name="rowCount">How many rows to build.</param>
        /// <param name="distinctKeys">Distinct values in the key column. Pass the row count for a
        /// key that is unique on every row (what a real export looks like); pass a small number to
        /// hold the shared-string table constant and isolate the streaming cost.</param>
        public static List<Row> Build(int rowCount, int distinctKeys)
        {
            if (distinctKeys < 1)
                throw new ArgumentOutOfRangeException(nameof(distinctKeys));

            var rows = new List<Row>(rowCount);
            for (var i = 0; i < rowCount; i++)
            {
                rows.Add(new Row
                {
                    LineNum = i + 1,
                    HeaderID = 100_000_000 + (i / 10),
                    DetailID = (i % 10) + 1,

                    // The key column: the only one whose distinct count the caller chooses.
                    Data = "D" + (i % distinctKeys).ToString(CultureInfo.InvariantCulture).PadLeft(9, '0'),

                    // Repeating columns. The strides are deliberately different so the
                    // combinations do not fall into a single repeating cycle.
                    Header01 = _branch[i % _branch.Length],
                    Header02 = _client[i % _client.Length],
                    Header03 = _product[i % _product.Length],
                    Header04 = _unit[i % _unit.Length],
                    Header05 = _state[i % _state.Length],
                    Header06 = _route[i % _route.Length],
                    Header07 = _staff[i % _staff.Length],
                    Footer01 = _kind1[i % _kind1.Length],
                    Footer02 = _kind2[i % _kind2.Length],
                    Footer03 = _kind3[i % _kind3.Length],
                    Footer04 = _kind4[i % _kind4.Length],
                    Footer05 = _kind5[i % _kind5.Length],
                    Footer06 = _kind6[i % _kind6.Length],
                    Footer07 = _kind7[i % _kind7.Length],
                    Footer08 = _kind8[i % _kind8.Length],
                });
            }
            return rows;
        }
    }
}
