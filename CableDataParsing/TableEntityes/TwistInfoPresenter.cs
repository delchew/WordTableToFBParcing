﻿using FirebirdDatabaseProvider.Attributes;

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

        [FBTableField(TableFieldName = "LAYER1_ELEMENTS_COUNT", TypeName = "INTEGER")]
        public int Layer1ElementsCount { get; set; }

        [FBTableField(TableFieldName = "LAYER2_ELEMENTS_COUNT", TypeName = "INTEGER")]
        public int Layer2ElementsCount { get; set; }

        [FBTableField(TableFieldName = "LAYER3_ELEMENTS_COUNT", TypeName = "INTEGER")]
        public int Layer3ElementsCount { get; set; }

        [FBTableField(TableFieldName = "LAYER4_ELEMENTS_COUNT", TypeName = "INTEGER")]
        public int Layer4ElementsCount { get; set; }

        [FBTableField(TableFieldName = "LAYER5_ELEMENTS_COUNT", TypeName = "INTEGER")]
        public int Layer5ElementsCount { get; set; }
    }
}
