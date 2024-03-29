using CableDataParsing.CableBulders;
using CablesDatabaseEFCoreFirebird.Entities;
using NUnit.Framework;
using Cables.Common;

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

            title = builder.GetCableTitle(cable, billet, CablePropertySet.HasArmourBraid);
            Assert.That(title, Is.EqualTo("КПСВВКГнг(А)-LS 8х2х1,0"));
        }

        [Test]
        public void KpsvvTitleTest2()
        {
            conductor.AreaInSqrMm = 2.5m;
            coverPolymerGroup.Title = "PE Self Extinguish";
            cable.ElementsCount = 4;

            title = builder.GetCableTitle(cable, billet, CablePropertySet.HasArmourBraid |
                                                 CablePropertySet.HasArmourTube |
                                                 CablePropertySet.HasFoilShield);
            Assert.That(title, Is.EqualTo("КПСВЭПсКПс 4х2х2,5"));
        }

        [Test]
        public void KpsvvTitleTest3()
        {
            conductor.AreaInSqrMm = 0.5m;
            coverPolymerGroup.Title = "PVC Cold";
            cable.ElementsCount = 24;

            title = builder.GetCableTitle(cable, billet, CablePropertySet.HasArmourTape |
                                                 CablePropertySet.HasArmourTube |
                                                 CablePropertySet.HasFoilShield);
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

        [Test]
        public void KpsvvTitleTest5()
        {
            conductor.AreaInSqrMm = 1m;
            coverPolymerGroup.Title = "PVC LSLTx";
            cable.ElementsCount = 9;

            title = builder.GetCableTitle(cable, billet, CablePropertySet.HasFoilShield);
            Assert.That(title, Is.EqualTo("ЛОУТОКС КПСВЭВнг(А)-LSLTx 9х2х1,0"));
        }
    }
}