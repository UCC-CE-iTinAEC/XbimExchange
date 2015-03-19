﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Xbim.COBieLiteUK;
using Xbim.DPoW;
using Attribute = Xbim.COBieLiteUK.Attribute;

namespace XbimExchanger.DPoWToCOBieLiteUK
{
    class MappingRoleToContact : MappingAttributableObjectToCOBieObject<Role, Xbim.COBieLiteUK.Contact>
    {
        protected override Xbim.COBieLiteUK.Contact Mapping(Role source, Xbim.COBieLiteUK.Contact target)
        {
            base.Mapping(source, target);

            target.ExternalId = source.Id.ToString();
            target.Category = "Role";
            target.GivenName = source.Name;

            //email has to be defined because it is a key for ContactKey references
            var email = String.Format("{0}@role.com", source.Name ?? "noname").ToLower();
            email = (new Regex("(\\s|\\[|\\])", RegexOptions.IgnoreCase)).Replace(email, "_").Trim('_').Trim();
            target.Email = email;

            if (target.Attributes == null) target.Attributes = new List<Attribute>();
            target.Attributes.Add("Name", "RoleName", source.Name, "RolePropertySet");
            target.Attributes.Add("Description", "RoleDescription", source.Description, "RolePropertySet");
            
            return target;
        }
    }
}
