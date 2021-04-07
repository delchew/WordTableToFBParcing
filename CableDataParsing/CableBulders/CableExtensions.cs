using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing.CableBulders
{
    public static class CableExtensions
    {
        public static Cable Clone (this Cable cable)
        {
            return new Cable
            {
                ClimaticMod = cable.ClimaticMod,
                ClimaticModId = cable.ClimaticModId,

                CoverColor = cable.CoverColor,
                CoverColorId = cable.CoverColorId,

                CoverPolymerGroup = cable.CoverPolymerGroup,
                CoverPolymerGroupId = cable.CoverPolymerGroupId,

                ElementsCount = cable.ElementsCount,

                FireProtectionClass = cable.FireProtectionClass,
                FireProtectionClassId = cable.FireProtectionClassId,

                FlatCableSize = cable.FlatCableSize,

                Id = 0,

                ListCableBillets = cable.ListCableBillets,
                
                ListCablePowerColors = cable.ListCablePowerColors,

                ListCableProperties = cable.ListCableProperties,

                MaxCoverDiameter = cable.MaxCoverDiameter,

                OperatingVoltage = cable.OperatingVoltage,
                OperatingVoltageId = cable.OperatingVoltageId,

                TechnicalConditions = cable.TechnicalConditions,
                TechnicalConditionsId = cable.TechnicalConditionsId,

                Title = cable.Title,

                TwistedElementType = cable.TwistedElementType,
                TwistedElementTypeId = cable.TwistedElementTypeId
            };
        }
    }
}
