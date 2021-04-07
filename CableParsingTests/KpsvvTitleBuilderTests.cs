using CableDataParsing.CableBulders;
using CablesDatabaseEFCoreFirebird.Entities;
using NUnit.Framework;

namespace CableParsingTests
{
    public class KpsvvTitleBuilderTests
    {
        string title;
        Conductor conductor;
        InsulatedBillet billet;
        Cable cable;
        PolymerGroup coverPolymerGroup;
        KpsvevTitleBuilder builder;

        [SetUp]
        public void Setup()
        {
            conductor = new Conductor();
            billet = new InsulatedBillet { Conductor = conductor };
            coverPolymerGroup = new PolymerGroup();
            builder = new KpsvevTitleBuilder();

            cable = new Cable
            {
                CoverPolymerGroup = coverPolymerGroup,
            };
        }

        [Test]
        public void KpsvvTitleTest1()
        {
            conductor.AreaInSqrMm = 1m;
            coverPolymerGroup.Title = "PVC LS";
            cable.ElementsCount = 8;

            title = builder.GetCableTitle(cable, billet, Cables.Common.CableProperty.HasArmourBraid);
            Assert.That(title, Is.EqualTo("КПСВВКГнг(А)-LS 8х2х1,0"));
        }

        [Test]
        public void KpsvvTitleTest2()
        {
            conductor.AreaInSqrMm = 2.5m;
            coverPolymerGroup.Title = "PE Self extinguish";
            cable.ElementsCount = 4;

            title = builder.GetCableTitle(cable, billet, Cables.Common.CableProperty.HasArmourBraid |
                                                 Cables.Common.CableProperty.HasArmourTube |
                                                 Cables.Common.CableProperty.HasFoilShield);
            Assert.That(title, Is.EqualTo("КПСВЭПсКПс 4х2х2,5"));
        }

        [Test]
        public void KpsvvTitleTest3()
        {
            conductor.AreaInSqrMm = 0.5m;
            coverPolymerGroup.Title = "PVC Cold";
            cable.ElementsCount = 24;

            title = builder.GetCableTitle(cable, billet, Cables.Common.CableProperty.HasArmourTape |
                                                 Cables.Common.CableProperty.HasArmourTube |
                                                 Cables.Common.CableProperty.HasFoilShield);
            Assert.That(title, Is.EqualTo("КПСВЭВБВм 24х2х0,5"));
        }

        [Test]
        public void KpsvvTitleTest4()
        {
            conductor.AreaInSqrMm = 1.5m;
            coverPolymerGroup.Title = "PVC Term";
            cable.ElementsCount = 12;

            title = builder.GetCableTitle(cable, billet, null);
            Assert.That(title, Is.EqualTo("КПСВВт 12х2х1,5"));
        }
    }
}