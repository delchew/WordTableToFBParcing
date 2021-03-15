using System.Collections.Generic;
using CableDataParsing;
using CablesDatabaseEFCoreFirebird.Entities;
using NUnit.Framework;

namespace CableParsingTests
{
    public class KpsvvParsingTests
    {
        CableParser parser;
        string title;
        Conductor conductor;
        InsulatedBillet billet;
        ListCableBillets billetsList;
        Cable cable;
        PolymerGroup coverPolymerGroup;

        [SetUp]
        public void Setup()
        {
            parser = new KpsvevParser(string.Empty, null);
            conductor = new Conductor();
            billet = new InsulatedBillet { Conductor = conductor };
            billetsList = new ListCableBillets { Billet = billet };
            coverPolymerGroup = new PolymerGroup();

            cable = new Cable
            {
                CoverPolymerGroup = coverPolymerGroup,
                ListCableBillets = new List<ListCableBillets> { billetsList },
            };
        }

        [Test]
        public void KpsvvTitleTest1()
        {
            conductor.AreaInSqrMm = 1m;
            coverPolymerGroup.Title = "PVC LS";
            cable.ElementsCount = 8;

            using (parser = new KpsvevParser(string.Empty, null))
                title = parser.GetCableTitle(cable, Cables.Common.CableProperty.HasArmourBraid);
            Assert.That(title, Is.EqualTo("КПСВВКГнг(А)-LS 8х2х1.0"));
        }

        [Test]
        public void KpsvvTitleTest2()
        {
            conductor.AreaInSqrMm = 2.5m;
            coverPolymerGroup.Title = "PE";
            cable.ElementsCount = 4;

            using (parser = new KpsvevParser(string.Empty, null))
                title = parser.GetCableTitle(cable, Cables.Common.CableProperty.HasArmourBraid |
                                                    Cables.Common.CableProperty.HasArmourTube |
                                                    Cables.Common.CableProperty.HasFoilShield);
            Assert.That(title, Is.EqualTo("КПСВЭПсКПс 4х2х2.5"));
        }

        [Test]
        public void KpsvvTitleTest3()
        {
            conductor.AreaInSqrMm = 0.5m;
            coverPolymerGroup.Title = "PVC Cold";
            cable.ElementsCount = 24;

            using (parser = new KpsvevParser(string.Empty, null))
                title = parser.GetCableTitle(cable, Cables.Common.CableProperty.HasArmourTape |
                                                    Cables.Common.CableProperty.HasArmourTube |
                                                    Cables.Common.CableProperty.HasFoilShield);
            Assert.That(title, Is.EqualTo("КПСВЭВБВм 24х2х0.5"));
        }
    }
}