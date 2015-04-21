﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xbim.COBieLiteUK;
using Xbim.DPoW;
using CAssetType = Xbim.COBieLiteUK.AssetType;
using DAssetType = Xbim.DPoW.AssetType;

namespace XbimExchanger.DPoWToCOBieLiteUK
{
    class MappingAssetTypeToAssetType: MappingDPoWObjectToCOBieObject<DAssetType, CAssetType>
    {
        protected override CAssetType Mapping(DAssetType sObject, CAssetType tObject)
        {
            //perform base mapping on the level of objects
            base.Mapping(sObject, tObject);

            //------------------ mappings specific to asset type ----------------------
            if (!String.IsNullOrEmpty(sObject.Variant))
            {
                //set variant of the object as a model number
                tObject.ModelNumber = sObject.Variant;
            }

            if (sObject.RequiredLOD != null)
                tObject.Shape = sObject.RequiredLOD.Code;

            //add dummy asset so that COBieLite => IFC will create real dummy geometry objects
            if (tObject.Assets == null) tObject.Assets = new List<Asset>();
            tObject.Assets.Add(new Asset
            {
                ExternalId = Guid.NewGuid().ToString(),
                ExternalSystem = "DPoW",
                Description = sObject.Description,
                Name = sObject.Name
            });

            // Mark asset types as required
            tObject.Categories = new List<Category> { new Category { Code = "required", Classification = "DPoW" } };

            return tObject;
        }
    }
}
