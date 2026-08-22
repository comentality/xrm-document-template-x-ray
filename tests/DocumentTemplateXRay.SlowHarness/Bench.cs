using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using DocumentTemplateXRay.Harness;
using ThreadExceptionEventHandler = System.Threading.ThreadExceptionEventHandler;

namespace DocumentTemplateXRay.SlowHarness
{
    /// <summary>One thing that happens, at a moment, to a tool somebody is using.</summary>
    public class Beat
    {
        public int At;
        public string Label;
        public Action<Run> Do;
    }

    /// <summary>
    /// A tool in a window, an environment that answers slowly, and a list of things that go wrong.
    /// Handed to every beat of a scenario.
    /// </summary>
    public class Run
    {
        public Form Form;
        public DocumentTemplateXRayControl Control;
        public Probe Probe;
        public SlowService Service;
        public string ShotDir;
        public Stall Stall;

        /// <summary>Everything the run has to answer for, findings and crashes alike.</summary>
        public readonly List<string> Failures = new List<string>();

        /// <summary>Captions of every dialog the tool put up, in order, closed as they appeared.</summary>
        public readonly List<string> Dialogs = new List<string>();

        public Stopwatch Clock = new Stopwatch();

        public void Check(bool ok, string what)
        {
            if (!ok) Failures.Add(Clock.ElapsedMilliseconds + "ms  " + what);
        }
    }

    /// <summary>
    /// How long the window went without being able to draw itself.
    ///
    /// Two ways of seeing the same thing. A gesture is timed directly: a click handler that does
    /// not return for a second is a window frozen for a second, and that is the measurement that
    /// catches work which should have been on a worker and was not. Between gestures a timer
    /// watches instead - it only ticks when the message loop is free, so the gap between two
    /// ticks is how long something else, a callback usually, held the thread.
    ///
    /// The harness pauses it around its own screenshots, because those block the loop too and
    /// are nobody's fault but the harness's.
    /// </summary>
    public class Stall
    {
        private readonly Timer _timer;
        private readonly Stopwatch _since = new Stopwatch();

        /// <summary>The longest the window was frozen, in milliseconds.</summary>
        public long Longest;

        /// <summary>What the tool was doing when that happened.</summary>
        public string Worst = "";

        public Stall()
        {
            _timer = new Timer { Interval = 25 };
            _timer.Tick += (s, e) =>
            {
                Saw(_since.ElapsedMilliseconds, "a callback landing");
                _since.Restart();
            };
        }

        public void Start() { _since.Restart(); _timer.Start(); }

        public void Stop() { _timer.Stop(); _timer.Dispose(); }

        /// <summary>Does one thing a person did, and notes how long the window was gone for.</summary>
        public void Gesture(string label, Action what)
        {
            var clock = Stopwatch.StartNew();
            try
            {
                what();
            }
            finally
            {
                clock.Stop();
                Saw(clock.ElapsedMilliseconds, label);
                _since.Restart();
            }
        }

        /// <summary>Does something the harness is responsible for, without blaming the tool.</summary>
        public void Ours(Action what)
        {
            _timer.Stop();
            try { what(); }
            finally { _since.Restart(); _timer.Start(); }
        }

        private void Saw(long waited, string doing)
        {
            if (waited <= Longest) return;
            Longest = waited;
            Worst = doing;
        }
    }

    public class Scenario
    {
        public string Name;

        /// <summary>One line, printed beside the result: which finding this is about.</summary>
        public string Why;

        /// <summary>How the environment behaves, set before anything is pressed.</summary>
        public Action<SlowService> Wire;

        public readonly List<Beat> Beats = new List<Beat>();

        public Scenario At(int ms, string label, Action<Run> what)
        {
            Beats.Add(new Beat { At = ms, Label = label, Do = what });
            return this;
        }

        /// <summary>
        /// Runs the whole scenario in a window of its own, and returns what went wrong.
        ///
        /// Everything happens on the UI thread with a live message loop, because that is where
        /// the tool lives: WorkAsync marshals its callback back here, and a scenario that drove
        /// the control from anywhere else would be testing something nobody does.
        /// </summary>
        public List<string> Play(string root)
        {
            var run = new Run { ShotDir = Path.Combine(root, Name), Stall = new Stall() };

            if (Directory.Exists(run.ShotDir)) Directory.Delete(run.ShotDir, true);
            Directory.CreateDirectory(run.ShotDir);

            run.Service = SlowService.Sampled();
            if (Wire != null) Wire(run.Service);

            run.Control = new DocumentTemplateXRayControl { Dock = DockStyle.Fill };
            Connect(run.Control, run.Service);
            run.Probe = new Probe(run.Control);

            run.Form = new QuietForm
            {
                Text = "Document Template XRay — " + Name,
                ClientSize = new Size(1280, 800)
            };
            run.Form.Controls.Add(run.Control);

            ThreadExceptionEventHandler onThreadException = (s, e) =>
                run.Failures.Add(run.Clock.ElapsedMilliseconds + "ms  the tool threw: " + Head(e.Exception));
            Application.ThreadException += onThreadException;

            var sweeper = new Dialogs(run);

            run.Form.Shown += (s, e) =>
            {
                run.Clock.Start();
                sweeper.Start();
                run.Stall.Start();

                var queue = new Queue<Beat>(Beats.OrderBy(b => b.At));
                var shot = new int[1];
                var beating = false;
                var pump = new Timer { Interval = 25 };
                pump.Tick += (s2, e2) =>
                {
                    // A beat that puts up a modal is still inside its own gesture, and a modal
                    // pumps messages - including this timer. Without the guard the next beat
                    // would fire from inside the last one, which is not a thing a person can do.
                    if (beating) return;
                    beating = true;
                    try
                    {
                        Due(run, queue, shot);
                    }
                    finally
                    {
                        beating = false;
                    }

                    if (queue.Count > 0) return;

                    pump.Stop();
                    sweeper.Stop();
                    run.Stall.Stop();
                    run.Form.Close();
                };
                pump.Start();
            };

            try
            {
                Application.Run(run.Form);
            }
            finally
            {
                Application.ThreadException -= onThreadException;
                sweeper.Stop();
                if (!run.Control.IsDisposed) run.Control.Dispose();
                run.Form.Dispose();
            }

            return run.Failures;
        }

        /// <summary>Every beat that has come due, in order, each with a shot of what it left.</summary>
        private static void Due(Run run, Queue<Beat> queue, int[] shot)
        {
            while (queue.Count > 0 && run.Clock.ElapsedMilliseconds >= queue.Peek().At)
            {
                var beat = queue.Dequeue();
                try
                {
                    run.Stall.Gesture(beat.Label, () => beat.Do(run));
                }
                catch (Exception ex)
                {
                    run.Failures.Add(run.Clock.ElapsedMilliseconds + "ms  \"" + beat.Label
                                     + "\" threw: " + Head(ex));
                }

                // A shot per beat, so a failure is something to look at rather than only a line
                // to read. Named in order, because the order is the story.
                run.Stall.Ours(() =>
                {
                    try
                    {
                        if (!run.Form.IsDisposed)
                        {
                            Capture.Of(run.Form, Path.Combine(run.ShotDir,
                                (++shot[0]).ToString("00") + "-" + Slug(beat.Label) + ".png"));
                        }
                    }
                    catch (Exception)
                    {
                        // A shot that cannot be taken is not worth failing a scenario over.
                    }
                });
            }
        }

        /// <summary>
        /// Hands the control its connection, the way XrmToolBox does. The setter is protected,
        /// because outside a host nothing has any business setting one - so a harness that is
        /// standing in for the host reaches it the way it reaches everything else here.
        /// </summary>
        private static void Connect(DocumentTemplateXRayControl control, Microsoft.Xrm.Sdk.IOrganizationService service)
        {
            const BindingFlags any = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
                                     | BindingFlags.DeclaredOnly;

            for (var type = control.GetType(); type != null; type = type.BaseType)
            {
                var property = type.GetProperty("Service", any);
                var setter = property == null ? null : property.GetSetMethod(true);
                if (setter != null)
                {
                    setter.Invoke(control, new object[] { service });
                    return;
                }

                var backing = type.GetField("<Service>k__BackingField", any);
                if (backing != null)
                {
                    backing.SetValue(control, service);
                    return;
                }
            }

            throw new MissingMemberException("PluginControlBase", "Service");
        }

        private static string Head(Exception ex)
        {
            var line = ex.GetBaseException().ToString();
            var stop = line.IndexOf('\n');
            return (stop < 0 ? line : line.Substring(0, stop)).Trim();
        }

        private static string Slug(string label)
        {
            var sb = new StringBuilder();
            foreach (var c in label ?? "beat")
            {
                sb.Append(char.IsLetterOrDigit(c) ? char.ToLowerInvariant(c) : '-');
            }

            return sb.ToString().Trim('-');
        }
    }

    /// <summary>
    /// Closes whatever modal the tool puts up, and writes down that it did.
    ///
    /// Nothing here is a workaround for the tool being awkward: a run with nobody at the keyboard
    /// has to answer a dialog or hang. The captions are worth keeping either way - "no dialog at
    /// all" is exactly the assertion a callback landing on a closed tab is caught by.
    /// </summary>
    internal class Dialogs
    {
        private const int WM_CLOSE = 0x0010;

        [DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(uint threadId, EnumThreadDelegate callback, IntPtr param);

        private delegate bool EnumThreadDelegate(IntPtr hwnd, IntPtr param);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern IntPtr PostMessage(IntPtr hwnd, int message, IntPtr wparam, IntPtr lparam);

        private readonly Run _run;
        private readonly Timer _timer;

        public Dialogs(Run run)
        {
            _run = run;
            // Faster than a person, slower than the message loop. A modal dialog runs its own
            // loop and pumps WM_TIMER, which is the whole reason this works from inside one.
            _timer = new Timer { Interval = 120 };
            _timer.Tick += (s, e) => Sweep();
        }

        public void Start() { _timer.Start(); }

        public void Stop() { _timer.Stop(); _timer.Dispose(); }

        private void Sweep()
        {
            var mine = _run.Form.IsDisposed ? IntPtr.Zero : _run.Form.Handle;
            EnumThreadWindows(GetCurrentThreadId(), (hwnd, param) =>
            {
                if (hwnd == mine || !IsWindowVisible(hwnd)) return true;

                var caption = new StringBuilder(260);
                GetWindowText(hwnd, caption, caption.Capacity);
                var text = caption.ToString();
                if (text.Length == 0) return true;

                _run.Dialogs.Add(text);
                PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
        }
    }
}
