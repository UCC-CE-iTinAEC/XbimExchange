using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xbim.DPoW;
using Xbim.COBieLiteUK;

namespace XbimExchanger.DPoWToCOBieLiteUK
{
    class MappingSpaceTypeToZone : MappingDPoWObjectToCOBieObject<SpaceType, Zone>
    {
        //implement any eventual specialities here
        protected override Xbim.COBieLiteUK.Zone Mapping(SpaceType source, Xbim.COBieLiteUK.Zone target)
        {
            base.Mapping(source, target);

            target.Categories = new List<Category> { new Category { Code = "required", Classification = "DPoW" } };

            return target;
        }
    }
}
