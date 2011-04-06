using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Class for generating slide title content
    /// </summary>
    class TitleBuilder
    {
        /// <summary>
        /// Number of current pass on slide (used for overlay)
        /// </summary>
        private int _passNumber;

        /// <summary>
        /// Discovered number of maximum passes (from overlay specification and pause commands)
        /// </summary>
        private int _maxPass = 1;

        /// <summary>
        /// Build (generate) title content
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="frametitle"></param>
        /// <param name="passNumber"></param>
        /// <returns>true if title is complete; false if needs another pass</returns>
        public bool BuildTitle(PowerPoint.Shape shape, FrametitleRecord frametitle, int passNumber)
        {
            _passNumber = passNumber;

            // todo: place code here

            return _passNumber == _maxPass;
        }
    }
}
