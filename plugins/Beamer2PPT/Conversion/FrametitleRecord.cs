using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Frame title record class
    /// </summary>
    public class FrametitleRecord
    {
        #region Private variables

        /// <summary>
        /// Title overlay specification
        /// </summary>
        private string _titleOverlaySpec;

        /// <summary>
        /// Expanded title overlay specification
        /// </summary>
        private ISet<int> _titleOverlaySet;

        /// <summary>
        /// Subtitle overlay specification
        /// </summary>
        private string _subtitleOverlaySpec;

        /// <summary>
        /// Expanded subtitle overlay specification
        /// </summary>
        private ISet<int> _subtitleOverlaySet;

        #endregion // Private variables

        /// <summary>
        /// Public constructor
        /// </summary>
        /// <param name="content">Content of frame title</param>
        /// <param name="content2">Content of frame subtitle</param>
        public FrametitleRecord(List<Node> title, List<Node> subtitle)
        {
            Title = title;
            Subtitle = subtitle;
        }

        #region Public properties

        /// <summary>
        /// Frame title content
        /// </summary>
        public List<Node> Title { get; set; }

        /// <summary>
        /// Frame title overlay setter
        /// </summary>
        public string TitleOverlay
        {
            set
            {
                _titleOverlaySpec = value;

                // invalidate set
                _titleOverlaySet = null;
            }
        }

        /// <summary>
        /// Frame title overlay specification getter
        /// </summary>
        public ISet<int> TitleOverlaySet
        {
            get
            {
                if (_titleOverlaySet == null)
                {
                    _titleOverlaySet = Misc.ParseOverlay(_titleOverlaySpec);
                }

                return _titleOverlaySet;
            }
        }

        /// <summary>
        /// Frame subtitle content
        /// </summary>
        public List<Node> Subtitle { get; set; }

        /// <summary>
        /// Frame subtitle overlay setter
        /// </summary>
        public string SubtitleOverlay
        {
            set
            {
                _subtitleOverlaySpec = value;

                // invalidate set
                _subtitleOverlaySet = null;
            }
        }

        /// <summary>
        /// Frame subtitle overlay specification getter
        /// </summary>
        public ISet<int> SubtitleOverlaySet
        {
            get
            {
                if (_subtitleOverlaySet == null)
                {
                    _subtitleOverlaySet = Misc.ParseOverlay(_subtitleOverlaySpec);
                }

                return _subtitleOverlaySet;
            }
        }

        #endregion // Public properties
    }
}
