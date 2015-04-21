using System;
using System.Linq;

namespace Xbim.DPoW
{
    /// <summary>
    /// Facility
    /// </summary>
    public class Facility
    {
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Description
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Category ID
        /// </summary>
        public Guid CategoryId { get; set; }
        /// <summary>
        /// Sita name
        /// </summary>
        public string SiteName { get; set; }
        /// <summary>
        /// Site description
        /// </summary>
        public string SiteDescription { get; set; }

        /// <summary>
        /// Gets category from actual plan of work by it's ID
        /// </summary>
        /// <param name="pow"></param>
        /// <returns></returns>
        public void GetCategory(PlanOfWork pow, out Classification classification, out ClassificationReference classificationReference)
        {
            classification = null;
            classificationReference = null;

            if (pow.ClassificationSystems == null) return;
            
            var result = (from c in pow.ClassificationSystems where c.ClassificationReferences != null from r in c.ClassificationReferences where r.Id == CategoryId select new { c, r }).FirstOrDefault();
            if (result == null) return;

            classification = result.c;
            classificationReference = result.r;
        }
    }
}
