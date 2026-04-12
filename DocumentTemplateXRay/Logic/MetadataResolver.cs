using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace DocumentTemplateXRay.Logic
{
    public class MetadataResolver
    {
        private readonly IOrganizationService _service;
        private readonly Dictionary<string, EntityMetadata> _cache = new Dictionary<string, EntityMetadata>(StringComparer.OrdinalIgnoreCase);

        public MetadataResolver(IOrganizationService service)
        {
            _service = service;
        }

        public void ResolveDisplayNames(List<FieldInfo> fields)
        {
            foreach (var field in fields)
            {
                if (string.IsNullOrEmpty(field.FieldPath)) continue;

                var segments = field.FieldPath.Split('/');
                if (segments.Length == 0) continue;

                // First segment is the entity, last segment is the attribute
                var entityName = segments[0];
                var attributeName = segments[segments.Length - 1];

                // For paths like entity/attribute, resolve directly
                // For paths like entity/relationship/attribute, the attribute belongs
                // to the related entity — we resolve what we can from the root entity
                var metadata = GetEntityMetadata(entityName);
                if (metadata == null) continue;

                if (segments.Length == 2)
                {
                    // Simple: entity/attribute
                    field.TableDisplayName = GetEntityDisplayName(metadata);
                    field.ColumnDisplayName = GetAttributeDisplayName(metadata, attributeName);
                }
                else
                {
                    // Relationship path: walk to the target entity for the last attribute
                    var targetMetadata = ResolveRelationshipPath(metadata, segments);
                    if (targetMetadata != null)
                    {
                        field.TableDisplayName = GetEntityDisplayName(targetMetadata);
                        field.ColumnDisplayName = GetAttributeDisplayName(targetMetadata, attributeName);
                    }
                }
            }
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

        private EntityMetadata GetEntityMetadata(string entityLogicalName)
        {
            if (_cache.TryGetValue(entityLogicalName, out var cached))
                return cached;

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
            catch
            {
                _cache[entityLogicalName] = null;
                return null;
            }
        }
    }
}
