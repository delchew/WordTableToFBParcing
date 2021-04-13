using CableDataParsing.CableBulders;
using CablesDatabaseEFCoreFirebird.Entities;
using NUnit.Framework;
using Cables.Common;

namespace CableParsingTests
{
    public class Kevv_KersTitleBuilderTests
    {
        string title;
        Conductor conductor;
        InsulatedBillet billet;
        Cable cable;
        PolymerGroup coverPolymerGroup;
        Kevv_KerspTitleBuilder builder;

        [SetUp]
        public void Setup()
        {
            conductor = new Conductor();
            billet = new InsulatedBillet { Conductor = conductor, PolymerGroup = new PolymerGroup() };
            coverPolymerGroup = new PolymerGroup();
            builder = new Kevv_KerspTitleBuilder();

            cable = new Cable
            {
                CoverPolymerGroup = coverPolymerGroup,
            };
        }

        [Test]
        public void KevvTitleTest()
        {
            conductor.AreaInSqrMm = 1m;
            coverPolymerGroup.Title = "PVC LS";
            billet.PolymerGroup.Title = "PVC LS";
            cable.ElementsCount = 8;

            title = builder.GetCableTitle(cable, billet, null);
            Assert.That(title, Is.EqualTo("КЭВВнг(А)-LS 8х1,0"));
        }

        [Test]
        public void KersepTitleTest()
        {
            conductor.AreaInSqrMm = 0.75m;
            coverPolymerGroup.Title = "HFCompound";
            billet.PolymerGroup.Title = "Rubber";
            cable.ElementsCount = 37;

            title = builder.GetCableTitle(cable, billet, CablePropertySet.HasBraidShield);
            Assert.That(title, Is.EqualTo("КЭРсЭПнг(А)-FRHF 37х0,75"));
        }


        [Test]
        public void KerseuTitleTest()
        {
            conductor.AreaInSqrMm = 0.5m;
            coverPolymerGroup.Title = "PUR";
            billet.PolymerGroup.Title = "Rubber";
            cable.ElementsCount = 2;

            title = builder.GetCableTitle(cable, billet, CablePropertySet.HasBraidShield);
            Assert.That(title, Is.EqualTo("КЭРсЭУнг(D)-FRHF 2х0,5"));
        }
    }
}
