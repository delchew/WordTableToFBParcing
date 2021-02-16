using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("TWIST_INFO")]
    public class TwistInfoPresenter
    {
        [FBTableField(TableFieldName = "ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "TWIST_INFO_ID_GEN")]
        public long Id { get; set; }

        [FBTableField(TableFieldName = "ELEMENTS_COUNT", TypeName = "INTEGER", IsNotNull = true)]
        public int ElementsCount { get; set; }

        [FBTableField(TableFieldName = "TWIST_KOEFFICIENT", TypeName = "NUMERIC(3, 2)")]
        public decimal TwistKoefficient { get; set; }

        [FBTableField(TableFieldName = "LAYERS_ELEMENTS_COUNT", TypeName = "INTEGER", IsNotNull = true)]
        public int [] LayersElementsCount { get; set; }
    }
}
