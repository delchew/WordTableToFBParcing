using CableDataParsing.CableBulders;
using CablesDatabaseEFCoreFirebird.Entities;
using NUnit.Framework;
using Cables.Common;

namespace CableParsingTests
{
    public class SkabTitleBuiderTests
    {
        string title;
        Conductor conductor;
        InsulatedBillet billet;
        Cable cable;
        PolymerGroup coverPolymerGroup;
        TwistedElementType twistedElementType;
        FireProtectionClass fireProtectionClass;
        SkabTitleBuilder builder;
        OperatingVoltage operatingVoltage250;
        OperatingVoltage operatingVoltage660;

        [SetUp]
        public void Setup()
        {
            conductor = new Conductor();
            billet = new InsulatedBillet { Conductor = conductor, PolymerGroup = new PolymerGroup() };
            coverPolymerGroup = new PolymerGroup();
            fireProtectionClass = new FireProtectionClass();
            twistedElementType = new TwistedElementType();
            operatingVoltage250 = new OperatingVoltage
            {
                ACVoltage = 380,
                ACFriquency = 400,
                DCVoltage = 540
            };
            operatingVoltage660 = new OperatingVoltage
            {
                ACVoltage = 660,
                ACFriquency = 400,
                DCVoltage = 1000
            };
            builder = new SkabTitleBuilder();

            cable = new Cable
            {
                CoverPolymerGroup = coverPolymerGroup,
                TwistedElementType = twistedElementType,
                FireProtectionClass = fireProtectionClass,
            };
        }

        [Test]
        public void SkabTitleTest1()
        {
            conductor.AreaInSqrMm = 1.5m;
            coverPolymerGroup.Title = "PVC LS";
            billet.PolymerGroup.Title = "PVC LS";
            cable.ElementsCount = 10;
            fireProtectionClass.Designation = "нг(А)-LS";
            twistedElementType.Id = 2;
            cable.OperatingVoltage = operatingVoltage250;

            var propertySet = CablePropertySet.HasBraidShield | CablePropertySet.HasFoilShield | CablePropertySet.HasFilling;

            title = builder.GetCableTitle(cable, billet, propertySet);
            Assert.That(title, Is.EqualTo("СКАБ 250нг(А)-LS 10х2х1,5л"));
        }

        [Test]
        public void SkabTitleTest2()
        {
            conductor.AreaInSqrMm = 1m;
            coverPolymerGroup.Title = "PVC LS";
            billet.PolymerGroup.Title = "Rubber";
            cable.ElementsCount = 7;
            fireProtectionClass.Designation = "нг(А)-FRLS";
            twistedElementType.Id = 3;
            cable.OperatingVoltage = operatingVoltage660;

            var propertySet =  CablePropertySet.HasFoilShield | CablePropertySet.HasIndividualFoilShields | CablePropertySet.HasArmourBraid;

            title = builder.GetCableTitle(cable, billet, propertySet);
            Assert.That(title, Is.EqualTo("СКАБ 660КГнг(А)-FRLS 7х3эх1,0л фо"));
        }

        [Test]
        public void SkabTitleTest3()
        {
            conductor.AreaInSqrMm = 0.75m;
            coverPolymerGroup.Title = "HFCompound";
            billet.PolymerGroup.Title = "HFCompound";
            cable.ElementsCount = 19;
            fireProtectionClass.Designation = "нг(А)-HF";
            twistedElementType.Id = 1;
            cable.OperatingVoltage = operatingVoltage250;

            var propertySet = CablePropertySet.HasBraidShield | CablePropertySet.HasFoilShield | CablePropertySet.HasFilling 
                                | CablePropertySet.SparkSafety | CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube
                                | CablePropertySet.HasWaterBlockStripe;

            title = builder.GetCableTitle(cable, billet, propertySet);
            Assert.That(title, Is.EqualTo("СКАБ 250Кнг(А)-HF-ХЛ 19х0,75л в Ex-i"));
        }

        [Test]
        public void SkabTitleTest4()
        {
            conductor.AreaInSqrMm = 2.5m;
            coverPolymerGroup.Title = "PUR";
            billet.PolymerGroup.Title = "Rubber";
            cable.ElementsCount = 24;
            fireProtectionClass.Designation = "нг(С)-FRHF";
            twistedElementType.Id = 2;
            cable.OperatingVoltage = operatingVoltage660;

            var propertySet = CablePropertySet.HasFoilShield | CablePropertySet.HasFilling
                                | CablePropertySet.HasArmourBraid | CablePropertySet.HasWaterBlockStripe;

            title = builder.GetCableTitle(cable, billet, propertySet);
            Assert.That(title, Is.EqualTo("СКАБ 660КГУнг(С)-FRHF-ХЛ 24х2х2,5л фв"));
        }
    }
}
