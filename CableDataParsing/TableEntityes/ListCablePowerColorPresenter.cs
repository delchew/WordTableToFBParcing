﻿using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("LIST_CABLE_POWER_COLOR")]
    public class ListCablePowerColorPresenter
    {
        [FBTableField(TableFieldName = "ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "LIST_CABLE_POWER_COLOR_ID_GEN")]
        public long ListId { get; set; }

        [FBTableField(TableFieldName = "CABLE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CableId { get; set; }

        [FBTableField(TableFieldName = "POWER_COLOR_SCHEME_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long PowerColorSchemeId { get; set; }
    }
}
