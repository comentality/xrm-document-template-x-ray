using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace DocumentTemplateXRay.Logic
{
    /// <summary>What one field path turned out to be, in words.</summary>
    public class ResolvedName
    {
        public string Table;
        public string Column;
    }

    public class MetadataResolver
    {
        private readonly IOrganizationService _service;
        private readonly Dictionary<string, EntityMetadata> _cache;

        /// <summary>
        /// Tables the environment could not be asked about at all - a timeout, a dropped
        /// connection. Not the same as a table that is not there, which is an answer.
        /// </summary>
        public readonly List<string> Unavailable = new List<string>();

        /// <summary>
        /// Everything looked up, including what was handed in. A RetrieveEntityRequest carries
        /// every attribute and every relationship of a table, so the caller keeps this and hands
        /// it back next time rather than paying for the same answer once per template.
        /// </summary>
        public Dictionary<string, EntityMetadata> Cache { get { return _cache; } }

        public MetadataResolver(IOrganizationService service,
            IDictionary<string, EntityMetadata> known = null)
        {
            _service = service;
            _cache = known == null
                ? new Dictionary<string, EntityMetadata>(StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, EntityMetadata>(known, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// The display names behind these field paths, one round trip per table not already known.
        ///
        /// It takes paths and gives back words, rather than taking the fields and writing into
        /// them: this runs on a worker, and those fields are being drawn on the UI thread at the
        /// same time. Nothing here touches anything the window owns.
        ///
        /// <paramref name="cancelled"/> is asked between tables. The request already on the wire
        /// cannot be recalled; the ones behind it are most of the wait on a slow link.
        /// </summary>
        public Dictionary<string, ResolvedName> Resolve(IEnumerable<string> fieldPaths,
            Func<bool> cancelled = null)
        {
            var resolved = new Dictionary<string, ResolvedName>(StringComparer.OrdinalIgnoreCase);

            foreach (var fieldPath in fieldPaths)
            {
                if (cancelled != null && cancelled()) break;
                if (string.IsNullOrEmpty(fieldPath) || resolved.ContainsKey(fieldPath)) continue;

                var segments = fieldPath.Split('/');
                if (segments.Length < 2) continue;

                // First segment is the entity, last segment is the attribute
                var entityName = segments[0];
                var attributeName = segments[segments.Length - 1];

                // For paths like entity/attribute, resolve directly
                // For paths like entity/relationship/attribute, the attribute belongs
                // to the related entity — we resolve what we can from the root entity
                var metadata = GetEntityMetadata(entityName);
                if (metadata == null) continue;

                // Relationship path: walk to the target entity for the last attribute
                var target = segments.Length == 2
                    ? metadata
                    : ResolveRelationshipPath(metadata, segments);
                if (target == null) continue;

                resolved[fieldPath] = new ResolvedName
                {
                    Table = GetEntityDisplayName(target),
                    Column = GetAttributeDisplayName(target, attributeName)
                };
            }

            return resolved;
        }

        private EntityMetadata ResolveRelationshipPath(EntityMetadata rootMetadata, string[] segments)
        {
            // segments: [entity, rel1, rel2, ..., attribute]
            // Walk through relationships to find the target entity for the last attribute
            var currentMetadata = rootMetadata;

            for (int i = 1; i < segments.Length - 1; i++)
            {
                var relName = segments[i];
                var targetEntity = FindRelationshipTarget(currentMetadata, relName);
                if (targetEntity == null) return null;

                currentMetadata = GetEntityMetadata(targetEntity);
                if (currentMetadata == null) return null;
            }

            return currentMetadata;
        }

        private string FindRelationshipTarget(EntityMetadata metadata, string relationshipSegment)
        {
            // Check one-to-many relationships
            if (metadata.OneToManyRelationships != null)
            {
                var rel = metadata.OneToManyRelationships.FirstOrDefault(r =>
                    string.Equals(r.SchemaName, relationshipSegment, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(r.ReferencingAttribute, relationshipSegment, StringComparison.OrdinalIgnoreCase));
                if (rel != null) return rel.ReferencingEntity;
            }

            // Check many-to-one relationships
            if (metadata.ManyToOneRelationships != null)
            {
                var rel = metadata.ManyToOneRelationships.FirstOrDefault(r =>
                    string.Equals(r.SchemaName, relationshipSegment, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(r.ReferencingAttribute, relationshipSegment, StringComparison.OrdinalIgnoreCase));
                if (rel != null) return rel.ReferencedEntity;
            }

            // Check many-to-many relationships
            if (metadata.ManyToManyRelationships != null)
            {
                var rel = metadata.ManyToManyRelationships.FirstOrDefault(r =>
                    string.Equals(r.SchemaName, relationshipSegment, StringComparison.OrdinalIgnoreCase));
                if (rel != null)
                {
                    return string.Equals(rel.Entity1LogicalName, metadata.LogicalName, StringComparison.OrdinalIgnoreCase)
                        ? rel.Entity2LogicalName
                        : rel.Entity1LogicalName;
                }
            }

            return null;
        }

        private string GetEntityDisplayName(EntityMetadata metadata)
        {
            return metadata?.DisplayName?.UserLocalizedLabel?.Label;
        }

        private string GetAttributeDisplayName(EntityMetadata metadata, string attributeLogicalName)
        {
            if (metadata.Attributes == null) return null;

            var attr = metadata.Attributes.FirstOrDefault(a =>
                string.Equals(a.LogicalName, attributeLogicalName, StringComparison.OrdinalIgnoreCase));

            return attr?.DisplayName?.UserLocalizedLabel?.Label;
        }

        /// <summary>
        /// Whether this is the link failing rather than the environment answering. A fault is an
        /// answer - that table is not there - and a timeout or a dropped channel is not.
        /// </summary>
        private static bool Unreachable(Exception error)
        {
            for (var ex = error; ex != null; ex = ex.InnerException)
            {
                if (ex is TimeoutException) return true;
                if (ex.GetType().Name.EndsWith("CommunicationException", StringComparison.Ordinal)) return true;
            }

            return false;
        }

        private EntityMetadata GetEntityMetadata(string entityLogicalName)
        {
            if (_cache.TryGetValue(entityLogicalName, out var cached))
                return cached;

            // Already tried this pass and the link was down. Not remembered past the pass - the
            // next attempt should ask again - but asking once per field of a template would be a
            // dozen timeouts for one answer nobody is going to get.
            if (Unavailable.Contains(entityLogicalName)) return null;

            try
            {
                var request = new RetrieveEntityRequest
                {
                    LogicalName = entityLogicalName,
                    EntityFilters = EntityFilters.Attributes | EntityFilters.Relationships
                };
                var response = (RetrieveEntityResponse)_service.Execute(request);
                _cache[entityLogicalName] = response.EntityMetadata;
                return response.EntityMetadata;
            }
            catch (Exception ex)
            {
                // A table the environment has no record of is an answer, and worth remembering.
                // A link that gave up is not an answer at all: remember nothing, so the next
                // attempt asks again, and say so, because the two look identical on screen.
                if (Unreachable(ex))
                {
                    if (!Unavailable.Contains(entityLogicalName)) Unavailable.Add(entityLogicalName);
                    return null;
                }

                _cache[entityLogicalName] = null;
                return null;
            }
        }
    }
}
