﻿/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */

namespace OpenXmlConverter
{
	/// <summary>
	/// Specifies the measurement values of a Html Unit.
	/// </summary>
    enum UnitMetric
	{
		Unknown,
		Percent,
		Inch,
		Centimeter,
		Millimeter,
		/// <summary>1em is equal to the current font size.</summary>
		EM,
		/// <summary>one ex is the x-height of a font (x-height is usually about half the font-size)</summary>
		Ex,
		Point,
		Pica,
		Pixel,

		// this value is not parsed but can be used internally
		Emus,

        /// <summary>Not convertible to any other units.</summary>
        Auto
	}
}