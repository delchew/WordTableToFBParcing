using GetInfoFromWordToFireBirdTable.CableEntityes;
using System;
using System.Text;

namespace GetInfoFromWordToFireBirdTable.Common.CableNameBuilders
{
    public class SkabNameBuilder : ICableNameBuilder<Skab> //TODO СДЕЛАТЬ!!!
    {
        private StringBuilder _nameBuilder;

        public SkabNameBuilder()
        {
            _nameBuilder = new StringBuilder("СКАБ ");
        }

        public string GetCableName(Skab cable)
        {
            return _nameBuilder.ToString();
        }
    }
}
