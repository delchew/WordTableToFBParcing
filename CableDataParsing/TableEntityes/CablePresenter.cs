﻿using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("CABLE")]
    public class CablePresenter
    {
        [FBTableField(TableFieldName = "ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "CABLE_ID_GEN")]
        public long Id { get; set; }

        [FBTableField(TableFieldName = "TITLE", TypeName = "VARCHAR(50)", IsNotNull = true)]
        public string Title { get; set; }

        [FBTableField(TableFieldName = "TWISTED_ELEMENT_TYPE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long TwistedElementTypeId { get; set; }

        [FBTableField(TableFieldName = "TECHNICAL_CONDITIONS_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long TechCondId { get; set; }

        [FBTableField(TableFieldName = "FIRE_PROTECT_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long FireProtectionId { get; set; }

        [FBTableField(TableFieldName = "COVER_POLYMER_GROUP_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CoverPolimerGroupId { get; set; }

        [FBTableField(TableFieldName = "COVER_COLOR_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CoverColorId { get; set; }

        [FBTableField(TableFieldName = "MAX_COVER_DIAMETER", TypeName = "NUMERIC(3, 2)")]
        public decimal? MaxCoverDiameter { get; set; }

        [FBTableField(TableFieldName = "CLIMATIC_MOD_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long ClimaticModId { get; set; }

        [FBTableField(TableFieldName = "OPERATING_VOLTAGE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long OperatingVoltageId { get; set; }

        [FBTableField(TableFieldName = "ELEMENTS_COUNT", TypeName = "NUMERIC(3, 1)", IsNotNull = true)]
        public decimal ElementsCount { get; set; }


    }
}