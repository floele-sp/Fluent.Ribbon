using System;
using System.Collections.Generic;
using System.Linq;

namespace Fluent
{
    using System.Collections.ObjectModel;
    using System.Windows;
    using System.Windows.Automation.Peers;
    using System.Windows.Automation.Provider;
    using System.Windows.Input;

    /// <summary>
    /// Contains all search results for a quick launch search.
    /// </summary>
    public class QuickLaunchResults
    {
        #region QuickLaunchResultsEventArgs

        /// <summary>
        /// Provides 
        /// </summary>
        public class QuickLaunchResultsEventArgs : EventArgs
        {
            /// <summary>
            /// Collection of search results, allows to add further results to the list.
            /// </summary>
            public QuickLaunchResults Results { get; set; }

            /// <summary>
            /// Current search subject.
            /// </summary>
            public string SearchSubject { get; set; }

            public QuickLaunchResultsEventArgs(QuickLaunchResults results, string searchSubject)
            {
                this.Results = results;
                this.SearchSubject = searchSubject;
            }
        }

        #endregion

        #region ResultItem

        /// <summary>
        /// Represents a single search result item.
        /// </summary>
        internal class ResultItem
        {
            public ICommandSource Target { get; set; }
            public string Text { get; set; }
            public IEnumerable<string> Path { get; set; }
            public object Icon { get; set; }
            public object CommandParameter { get; set; }
            public IInvokeProvider InvokeProvider { get; set; }
            public Action Action { get; set; }
            public string Group { get; set; }

            public int PathLength
            {
                get
                {
                    return this.Path.Count();
                }
            }

            public string DisplayPath
            {
                get
                {
                    return string.Join(" → ", this.Path.ToArray()).Trim(' ', '→') + " → ";
                }
            }

            public void Execute(IInputElement target)
            {
                if (this.Action != null)
                {
                    this.Action();
                }
                else if (this.Target.Command != null)
                {
                    RoutedCommand routedCommand = this.Target.Command as RoutedCommand;
                    if (routedCommand != null)
                    {
                        routedCommand.Execute(CommandParameter, target);
                    }
                    else
                    {
                        this.Target.Command.Execute(CommandParameter);
                    }
                }
                else if (this.InvokeProvider != null)
                {
                    this.InvokeProvider.Invoke();
                }
            }

            public override string ToString()
            {
                return this.Text;
            }
        }

        #endregion

        private ObservableCollection<ResultItem> searchResults = new ObservableCollection<ResultItem>();

        #region Properties

        /// <summary>
        /// Gets all search results.
        /// </summary>
        internal ObservableCollection<ResultItem> SearchResults
        {
            get { return this.searchResults; }
        }

        #endregion

        /// <summary>
        /// Adds a custom item to the search result.
        /// </summary>
        public void Add(string displayText, IEnumerable<string> path, string group, Action action)
        {
            this.searchResults.Add(new ResultItem() { Text = displayText, Path = path, Action = action, Group = group });
        }

        /// <summary>
        /// Adds a search result based on a button.
        /// </summary>
        internal void Add(Button button, IEnumerable<string> path, object fallbackIcon)
        {
            if (button.IsEnabled && button.Command != null)
            {
                this.searchResults.Add(new ResultItem()
                {
                    Text = button.Header as string,
                    Target = button,
                    Icon = button.Icon ?? fallbackIcon,
                    Path = path,
                    CommandParameter = button.CommandParameter,
                    Group = Ribbon.Localization.QuickLaunchMenuGroup
                });
            }
        }

        /// <summary>
        /// Adds a search result based on a button.
        /// </summary>
        internal void Add(ToggleButton button, IEnumerable<string> path, object fallbackIcon)
        {
            if (button.IsEnabled && button.Command != null)
            {
                this.searchResults.Add(new ResultItem()
                {
                    Text = button.Header as string,
                    Target = button,
                    Icon = button.Icon ?? fallbackIcon,
                    Path = path,
                    CommandParameter = button.CommandParameter,
                    Group = Ribbon.Localization.QuickLaunchMenuGroup
                });
            }
        }

        /// <summary>
        /// Adds a search result based on a split button.
        /// </summary>
        internal void Add(SplitButton splitButton, IEnumerable<string> path, object fallbackIcon)
        {
            if (splitButton.IsEnabled && splitButton.Command != null)
            {
                this.searchResults.Add(new ResultItem()
                {
                    Text = splitButton.Header as string,
                    Target = splitButton,
                    Icon = splitButton.Icon ?? fallbackIcon,
                    Path = path,
                    CommandParameter =
                    splitButton.CommandParameter,
                    Group = Ribbon.Localization.QuickLaunchMenuGroup
                });
            }
        }

        /// <summary>
        /// Adds a search result based on a menu item.
        /// </summary>
        internal void Add(MenuItem item, IEnumerable<string> path, object fallbackIcon)
        {
            if (item.IsEnabled)
            {
                MenuItemAutomationPeer peer = new MenuItemAutomationPeer(item);
                IInvokeProvider invokeProv = peer.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                this.searchResults.Add(new ResultItem()
                {
                    Text = item.Header as string,
                    Target = item,
                    Icon = item.Icon ?? fallbackIcon,
                    Path = path,
                    CommandParameter = item.CommandParameter,
                    InvokeProvider = invokeProv,
                    Group = Ribbon.Localization.QuickLaunchMenuGroup
                });
            }
        }
    }
}
