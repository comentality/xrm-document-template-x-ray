using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using DocumentTemplateXRay.Harness;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace DocumentTemplateXRay.SlowHarness
{
    /// <summary>One round trip, as it happened.</summary>
    public class Call
    {
        /// <summary>Position among every call of the run, from 1.</summary>
        public int Index;

        /// <summary>Position among the calls of this kind, from 1. What a scenario scripts against.</summary>
        public int Nth;

        /// <summary>templates, or entity.</summary>
        public string What;

        /// <summary>Which table the metadata question was about.</summary>
        public string Detail;

        public DateTime Started;
        public DateTime Ended;

        public override string ToString()
        {
            return "#" + Index + " " + What
                   + (string.IsNullOrEmpty(Detail) ? "" : " (" + Detail + ")")
                   + " " + (int)(Ended - Started).TotalMilliseconds + "ms";
        }
    }

    /// <summary>
    /// The environment, at a distance. Answers the two things this tool asks - every Word
    /// document template with its content, and one table's metadata at a time - out of
    /// <see cref="Sample"/>, and takes as long over each as the scenario says it should.
    ///
    /// The second one is the interesting one. Display-name resolution walks a field path table by
    /// table, and each hop is a <c>RetrieveEntityRequest</c> carrying every attribute and every
    /// relationship of that table. One template is several of them, one after another, which on a
    /// slow link is where the whole wait lives - and is why they are logged one by one.
    /// </summary>
    public class SlowService : IOrganizationService
    {
        private readonly object _lock = new object();
        private readonly List<Call> _calls = new List<Call>();
        private int _index;
        private readonly Dictionary<string, int> _perKind = new Dictionary<string, int>();

        public readonly List<SampleTemplate> Templates;
        public readonly Dictionary<string, SampleEntity> Tables;

        /// <summary>How long this call should take. The whole point of the harness.</summary>
        public Func<Call, int> Latency = call => 0;

        /// <summary>What this call should throw instead of answering, or null to answer.</summary>
        public Func<Call, Exception> Fails = call => null;

        public SlowService(List<SampleTemplate> templates, Dictionary<string, SampleEntity> tables)
        {
            Templates = templates;
            Tables = tables;
        }

        /// <summary>The sample environment: the fixture templates, and the tables they name.</summary>
        public static SlowService Sampled()
        {
            return new SlowService(Sample.Templates(), Sample.Metadata());
        }

        public List<Call> Log()
        {
            lock (_lock) return _calls.ToList();
        }

        public List<Call> Log(string what)
        {
            return Log().Where(c => c.What == what).ToList();
        }

        /// <summary>Every table this environment was asked about, in order, repeats included.</summary>
        public List<string> TablesAsked()
        {
            return Log("entity").Select(c => c.Detail).ToList();
        }

        /// <summary>
        /// Whether two calls ever overlapped in time. A tool that asks one question at a time
        /// puts one round trip on the wire at a time, and this is how that is checked rather
        /// than assumed.
        /// </summary>
        public bool Overlapped()
        {
            var log = Log().OrderBy(c => c.Started).ToList();
            for (var i = 1; i < log.Count; i++)
            {
                if (log[i].Started < log[i - 1].Ended) return true;
            }

            return false;
        }

        private Call Begin(string what, string detail)
        {
            var call = new Call { What = what, Detail = detail, Started = DateTime.UtcNow };

            lock (_lock)
            {
                call.Index = ++_index;
                int nth;
                _perKind.TryGetValue(what, out nth);
                _perKind[what] = call.Nth = nth + 1;
                _calls.Add(call);
            }

            return call;
        }

        /// <summary>
        /// The answer as it stood when the question arrived, and then the wait.
        ///
        /// That order is load bearing rather than tidy. A query is settled where the data is, and
        /// only the answer travels; a fake that slept first and read afterwards would hand a
        /// question asked three seconds ago the environment as it is now. A failure waits just as
        /// long: a link that is going to time out takes as long about it as one that works.
        /// </summary>
        private T Answer<T>(Call call, Func<T> answer)
        {
            var failure = Fails(call);
            var result = failure == null ? answer() : default(T);

            var delay = Latency(call);
            if (delay > 0) Thread.Sleep(delay);

            lock (_lock) call.Ended = DateTime.UtcNow;
            if (failure != null) throw failure;
            return result;
        }

        public EntityCollection RetrieveMultiple(QueryBase query)
        {
            var q = query as QueryExpression;
            if (q == null) throw new NotSupportedException("The tool only issues QueryExpressions.");
            if (q.EntityName != "documenttemplate")
                throw new NotSupportedException("Nothing in the tool queries " + q.EntityName + ".");

            return Answer(Begin("templates", null), () =>
            {
                lock (_lock)
                {
                    var rows = Templates.Select(t =>
                    {
                        var e = new Entity("documenttemplate", Guid.NewGuid());
                        e["name"] = t.Name;
                        e["documenttype"] = new OptionSetValue(2);
                        e["associatedentitytypecode"] = t.EntityType;
                        e["content"] = Content(t.Path);
                        return e;
                    }).ToList();

                    return new EntityCollection(rows) { MoreRecords = false };
                }
            });
        }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            var retrieve = request as RetrieveEntityRequest;
            if (retrieve == null)
                throw new NotSupportedException("Nothing in the tool executes " + request.RequestName + ".");

            var name = retrieve.LogicalName;
            return Answer(Begin("entity", name), () =>
            {
                SampleEntity table;
                lock (_lock)
                {
                    if (!Tables.TryGetValue(name, out table))
                    {
                        // What a real environment does about a table that is not there, and what
                        // fixture 04 is pointed at.
                        throw new InvalidOperationException(
                            "Could not find an entity with the name " + name + ".");
                    }
                }

                var response = new RetrieveEntityResponse();
                response.Results = new ParameterCollection
                {
                    { "EntityMetadata", Metadata(table) }
                };
                return (OrganizationResponse)response;
            });
        }

        /// <summary>The file, as the environment stores it: base64 of the bytes on disk.</summary>
        private static readonly Dictionary<string, string> Encoded =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private static string Content(string path)
        {
            string encoded;
            if (Encoded.TryGetValue(path, out encoded)) return encoded;
            encoded = Convert.ToBase64String(File.ReadAllBytes(path));
            Encoded[path] = encoded;
            return encoded;
        }

        /// <summary>
        /// One table, as metadata. Most of what the resolver reads is declared without a public
        /// setter, because outside a real environment nothing has any business inventing metadata
        /// - so a harness standing in for one reaches it the way it reaches everything else here.
        /// </summary>
        private EntityMetadata Metadata(SampleEntity table)
        {
            var metadata = new EntityMetadata();
            Set(metadata, "LogicalName", table.LogicalName);
            Set(metadata, "DisplayName", Label(table.DisplayName));

            var attributes = table.Attributes.Select(pair =>
            {
                var attribute = new StringAttributeMetadata();
                Set(attribute, "LogicalName", pair.Key);
                Set(attribute, "DisplayName", Label(pair.Value));
                return (AttributeMetadata)attribute;
            }).ToArray();
            Set(metadata, "Attributes", attributes);

            Set(metadata, "ManyToOneRelationships", table.ManyToOne.Select(pair =>
            {
                var relationship = new OneToManyRelationshipMetadata();
                Set(relationship, "SchemaName", pair.Key);
                Set(relationship, "ReferencingEntity", table.LogicalName);
                Set(relationship, "ReferencedEntity", pair.Value);
                return relationship;
            }).ToArray());

            Set(metadata, "OneToManyRelationships", table.OneToMany.Select(pair =>
            {
                var relationship = new OneToManyRelationshipMetadata();
                Set(relationship, "SchemaName", pair.Key);
                Set(relationship, "ReferencingEntity", pair.Value);
                Set(relationship, "ReferencedEntity", table.LogicalName);
                return relationship;
            }).ToArray());

            Set(metadata, "ManyToManyRelationships", new ManyToManyRelationshipMetadata[0]);

            return metadata;
        }

        private static Microsoft.Xrm.Sdk.Label Label(string text)
        {
            var localized = new LocalizedLabel(text, 1033);
            var label = new Microsoft.Xrm.Sdk.Label(localized, new[] { localized });
            if (label.UserLocalizedLabel == null) Set(label, "UserLocalizedLabel", localized);
            return label;
        }

        private static void Set(object target, string name, object value)
        {
            const BindingFlags any = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            for (var type = target.GetType(); type != null; type = type.BaseType)
            {
                var property = type.GetProperty(name, any | BindingFlags.DeclaredOnly);
                var setter = property == null ? null : property.GetSetMethod(true);
                if (setter != null)
                {
                    setter.Invoke(target, new[] { value });
                    return;
                }

                var backing = type.GetField("<" + name + ">k__BackingField", any)
                              ?? type.GetField("_" + char.ToLowerInvariant(name[0]) + name.Substring(1), any);
                if (backing != null)
                {
                    backing.SetValue(target, value);
                    return;
                }
            }

            throw new MissingMemberException(target.GetType().Name, name);
        }

        public Guid Create(Entity entity) { throw new NotSupportedException("The tool never writes."); }
        public void Update(Entity entity) { throw new NotSupportedException("The tool never writes."); }
        public void Delete(string entityName, Guid id) { throw new NotSupportedException("The tool never writes."); }
        public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { throw new NotSupportedException("The tool never writes."); }
        public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities) { throw new NotSupportedException("The tool never writes."); }
        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) { throw new NotSupportedException("The tool retrieves in bulk only."); }
    }
}
