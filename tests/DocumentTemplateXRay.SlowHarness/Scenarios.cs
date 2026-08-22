using System;
using System.Collections.Generic;
using System.Linq;
using DocumentTemplateXRay.Harness;

namespace DocumentTemplateXRay.SlowHarness
{
    /// <summary>
    /// One scenario per thing a slow link does to this tool. Each is a sequence of ordinary
    /// gestures at ordinary intervals, against an environment that takes seconds rather than
    /// milliseconds, and a handful of questions asked of the window in between.
    ///
    /// None of them is about which fields the tool finds in a template - the fixtures are for
    /// that. They are about the window while the answer is still on its way: which buttons a
    /// person can still press, whether the list and the pane are describing the same template,
    /// and whether the tool goes on asking about a template nobody is looking at any more.
    /// </summary>
    public static class Scenarios
    {
        private const string Summary = "Account Summary";
        private const string Letterhead = "Account Letterhead";
        private const string Blank = "Blank Letter";
        private const string Big = "Big Letterhead";

        /// <summary>
        /// A link where one kind of question is slow and the rest are merely remote. Metadata
        /// latency is per call, because display-name resolution is several calls one after
        /// another and the point of most of this is what happens between them.
        /// </summary>
        private static Func<Call, int> Slow(string what, params int[] perCall)
        {
            return call =>
            {
                if (call.What != what) return 20;
                return perCall[Math.Min(call.Nth - 1, perCall.Length - 1)];
            };
        }

        public static List<Scenario> All()
        {
            return new List<Scenario>
            {
                CloseDuringFetch(),
                FetchTwice(),
                FetchClobbersSelection(),
                ResolveAbandoned(),
                MetadataReasked(),
                FetchFails(),
                MetadataFails(),
                Cancel(),
                BigTemplate(),
            };
        }

        /// <summary>
        /// Somebody gives up and closes the tab while the environment is still thinking. The
        /// answer arrives afterwards and is handed to a control that no longer exists - and puts
        /// a modal error up over whatever the person moved on to, about a tool that is not on
        /// screen any more.
        /// </summary>
        private static Scenario CloseDuringFetch()
        {
            return new Scenario
            {
                Name = "close-during-fetch",
                Why = "a fetch landing on a closed tab must neither throw nor say anything",
                Wire = s =>
                {
                    s.Latency = Slow("templates", 3000);
                    s.Fails = call => call.What == "templates"
                        ? new TimeoutException("The request channel timed out while waiting for a reply.")
                        : null;
                }
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(1000, "close the tab", r =>
                {
                    r.Form.Controls.Remove(r.Control);
                    r.Control.Dispose();
                })
                .At(5000, "let the answer land on nobody", r =>
                {
                    r.Check(r.Service.Log("templates").Count == 1,
                        "the fetch should have gone out exactly once");
                    r.Check(r.Dialogs.Count == 0,
                        "a closed tab must not put a dialog up afterwards, and put up: "
                        + string.Join(" | ", r.Dialogs));
                });
        }

        /// <summary>
        /// Fetch pressed again because the first press did not appear to do anything. Every
        /// document template in the environment comes back with its content attached - the whole
        /// Word file, per template - so this is the most expensive request the tool makes, and it
        /// used to be possible to have three of them in the air at once.
        /// </summary>
        private static Scenario FetchTwice()
        {
            return new Scenario
            {
                Name = "fetch-twice",
                Why = "Fetch must be dead while it is fetching",
                Wire = s => s.Latency = Slow("templates", 2500)
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(800, "press it again", r =>
                {
                    r.Check(!r.Probe.Fetch.Enabled, "Fetch should be dead while it is fetching");
                    r.Probe.PressFetch();
                })
                .At(1600, "and again", r => r.Probe.PressFetch())
                .At(5000, "count the round trips", r =>
                {
                    r.Check(r.Service.Log("templates").Count == 1,
                        "three presses, one query, and there were " + r.Service.Log("templates").Count);
                    r.Check(r.Probe.Fetch.Enabled, "and Fetch should be live again afterwards");
                });
        }

        /// <summary>
        /// A template is open and being read when a refresh lands. The list is rebuilt from
        /// scratch, so the selection goes - and the pane goes on showing the fields of a template
        /// that is no longer selected, no longer highlighted, and not obviously still in the list.
        /// </summary>
        private static Scenario FetchClobbersSelection()
        {
            return new Scenario
            {
                Name = "fetch-clobbers-selection",
                Why = "a refresh must not leave the list and the pane describing different things",
                Wire = s => s.Latency = Slow("templates", 200, 2500)
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(1200, "open a template", r => r.Probe.SelectTemplate(Summary))
                .At(1800, "which is what is on screen", r =>
                {
                    r.Check(r.Probe.SelectedTemplate() != null, "the template should be selected");
                    r.Check(r.Probe.ResultsShowing, "and its fields showing");
                })
                .At(2000, "press Fetch again", r => r.Probe.PressFetch())
                .At(5000, "once the refresh lands", r =>
                {
                    r.Check(r.Probe.SelectedTemplate() != null,
                        "the template being read should still be the selected one");
                    r.Check(r.Probe.FilePath.IndexOf(Summary, StringComparison.Ordinal) >= 0,
                        "and the pane should still name it, and names \"" + r.Probe.FilePath + "\"");
                })
                .At(6400, "and settles", r =>
                    r.Check(r.Probe.FieldPaths().Count > 0,
                        "with its fields still on screen"));
        }

        /// <summary>
        /// Two templates opened one after another, which is what browsing a list is. Resolving
        /// display names walks the field paths table by table, a round trip each, and the ones
        /// left over belong to a template nobody is looking at any more.
        ///
        /// The request on the wire cannot be recalled. The ones behind it can, and on a slow link
        /// those are most of the wait.
        /// </summary>
        private static Scenario ResolveAbandoned()
        {
            return new Scenario
            {
                Name = "resolve-abandoned",
                Why = "moving on must stop the round trips the old template had left",
                Wire = s => s.Latency = Slow("entity", 1200)
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(1200, "open a template with three tables in it", r => r.Probe.SelectTemplate(Summary))
                .At(2900, "open one with none", r => r.Probe.SelectTemplate(Blank))
                .At(7000, "count the round trips", r =>
                {
                    var asked = r.Service.TablesAsked();
                    r.Check(asked.Count <= 2,
                        "only the request already on the wire should have finished, and there were "
                        + asked.Count + ": " + string.Join(", ", asked));
                    r.Check(r.Probe.FieldPaths().Count == 0,
                        "and the pane should belong to the template opened last");
                });
        }

        /// <summary>
        /// Three templates on the same four tables. The resolver caches what it has looked up, but
        /// the cache is built inside the resolver and thrown away with it - so every template, and
        /// every reopening of the same template, asks the environment for the same metadata again.
        ///
        /// A RetrieveEntityRequest carries every attribute and every relationship of a table. It
        /// is not a request to make twice for the same answer, let alone nine times.
        /// </summary>
        private static Scenario MetadataReasked()
        {
            return new Scenario
            {
                Name = "metadata-reasked",
                Why = "a table already looked up must not be looked up again",
                Wire = s => s.Latency = Slow("entity", 200)
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(1200, "open a template", r => r.Probe.SelectTemplate(Summary))
                .At(2900, "open another on the same tables", r => r.Probe.SelectTemplate(Letterhead))
                .At(4600, "go back to the first", r => r.Probe.SelectTemplate(Summary))
                .At(6500, "count the round trips", r =>
                {
                    var asked = r.Service.TablesAsked();
                    r.Check(asked.Count == asked.Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                        "no table should be asked about twice, and it asked for: "
                        + string.Join(", ", asked));
                    r.Check(r.Probe.ResolvedPaths().Count > 0,
                        "and the names should still be resolved from the cache");
                });
        }

        /// <summary>
        /// The link gives up rather than answering. A dialog is read once and dismissed; what is
        /// on screen afterwards is an empty template list, which reads as an environment with no
        /// Word templates in it. And if the flag that holds Fetch down does not come up on the
        /// failure path, that is a tool bricked by one timeout.
        /// </summary>
        private static Scenario FetchFails()
        {
            return new Scenario
            {
                Name = "fetch-fails",
                Why = "a fetch that fails must say so, and must give the tool back",
                Wire = s =>
                {
                    s.Latency = Slow("templates", 1500, 200);
                    s.Fails = call => call.What == "templates" && call.Nth == 1
                        ? new TimeoutException("The request channel timed out while waiting for a reply.")
                        : null;
                }
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(3000, "the tool is still a tool", r =>
                {
                    r.Check(r.Probe.Fetch.Enabled, "Fetch must come back, or there is no way on from here");
                    r.Check(r.Probe.DropText.IndexOf("could not", StringComparison.OrdinalIgnoreCase) >= 0,
                        "and the pane should say why the list is empty, and says \""
                        + r.Probe.DropText.Replace("\n", " ") + "\"");
                })
                .At(3200, "press Fetch again", r => r.Probe.PressFetch())
                .At(5000, "and this time it lands", r =>
                {
                    r.Check(r.Probe.TemplateNames().Count > 0, "the second fetch should fill the list");
                    r.Check(r.Probe.DropText.IndexOf("could not", StringComparison.OrdinalIgnoreCase) < 0,
                        "with the failure gone from the pane");
                });
        }

        /// <summary>
        /// The metadata cannot be read at all. The tool shows the field paths anyway, with the
        /// Table and Column cells empty - which is exactly what it shows for a column that has no
        /// display name, and for one that was deleted from the environment.
        ///
        /// This is the tool's whole point: telling apart three fields that all read "description".
        /// A blank cell that might mean "could not ask" is not an answer to that.
        /// </summary>
        private static Scenario MetadataFails()
        {
            return new Scenario
            {
                Name = "metadata-fails",
                Why = "unread names must not be passed off as names that are not there",
                Wire = s =>
                {
                    s.Latency = Slow("entity", 400);
                    s.Fails = call => call.What == "entity"
                        ? new TimeoutException("The request channel timed out while waiting for a reply.")
                        : null;
                }
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(1200, "open a template", r => r.Probe.SelectTemplate(Summary))
                .At(3900, "read what it says", r =>
                {
                    r.Check(r.Probe.FieldPaths().Count > 0,
                        "the fields themselves come out of the file, so they should be listed");
                    r.Check(r.Probe.FieldCount.IndexOf("could not", StringComparison.OrdinalIgnoreCase) >= 0
                            || r.Probe.FieldCount.IndexOf("failed", StringComparison.OrdinalIgnoreCase) >= 0,
                        "but the names could not be read, and the tool says \"" + r.Probe.FieldCount + "\"");
                });
        }

        /// <summary>
        /// The panel's Cancel button, which the tool never offered because no fetch was marked
        /// cancelable - so pressing it did nothing at best, and threw at worst.
        /// </summary>
        private static Scenario Cancel()
        {
            return new Scenario
            {
                Name = "cancel",
                Why = "a resolve must be abandonable, and must leave nothing stuck behind",
                Wire = s => s.Latency = Slow("entity", 1500)
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(1200, "open a template", r => r.Probe.SelectTemplate(Summary))
                .At(1900, "give up", r => r.Probe.PressCancel())
                .At(5500, "nothing more went out", r =>
                {
                    var asked = r.Service.TablesAsked();
                    r.Check(asked.Count == 1,
                        "cancelling should stop the round trips that had not happened yet, and there were "
                        + asked.Count + ": " + string.Join(", ", asked));
                    r.Check(r.Probe.FieldCount.IndexOf("cancel", StringComparison.OrdinalIgnoreCase) >= 0,
                        "and should say so rather than looking like a template with no names: \""
                        + r.Probe.FieldCount + "\"");
                    r.Check(r.Probe.Fetch.Enabled, "the tool should be usable again");
                });
        }

        /// <summary>
        /// A template that is big rather than complicated - a letterhead with artwork and
        /// boilerplate, which is most of them. Decoding it, writing it to a temp file, unzipping
        /// it and scanning it all happened on the thread that draws the window, so the tool went
        /// white for as long as the file took.
        ///
        /// The file here is local and fast; a template on a share that is having a bad day is the
        /// same code taking as long as the share does.
        /// </summary>
        private static Scenario BigTemplate()
        {
            return new Scenario
            {
                Name = "big-template",
                Why = "reading a template must not freeze the window",
                Wire = s =>
                {
                    s.Latency = Slow("templates", 200);
                    s.Templates.Add(new SampleTemplate
                    {
                        Name = Big,
                        EntityType = "account",
                        Path = Sample.BigTemplate()
                    });
                }
            }
                .At(200, "press Fetch", r => r.Probe.PressFetch())
                .At(1200, "open the big one", r => r.Probe.SelectTemplate(Big))
                .At(5000, "how long was the window frozen", r =>
                {
                    r.Check(r.Stall.Longest < 150,
                        "the window should stay drawable, and was frozen for " + r.Stall.Longest
                        + "ms during \"" + r.Stall.Worst + "\"");
                    r.Check(r.Probe.FieldPaths().Count == 4,
                        "and the fields should be there, and there were "
                        + r.Probe.FieldPaths().Count);
                });
        }
    }
}
