// Copyright (c) 2014-2017 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
// Copyright (c) 2010-2018 Andrii Tkachenko
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class sheet_drawings;

namespace detail
{
	class xlsx_consumer;
	class xlsx_producer;
}

/// <summary>
/// </summary>
class XLNT_API sheet_drawing
{
public:
	sheet_drawing()
	{
		from_col_offset = 0;
		from_row_offset = 0;
	}

	cell_reference from;
	int from_col_offset;
	int from_row_offset;
	optional<cell_reference> to;
	optional<int> to_col_offset;
	optional<int> to_row_offset;
	optional<std::string> get_picture_name() const
	{
		return picture_name;
	}
private:
	friend class sheet_drawings;
	friend class detail::xlsx_consumer;
	friend class detail::xlsx_producer;
	optional<int> picture_id;
	optional<std::string> picture_rel;
	optional<std::string> picture_name;
};

} // namespace xlnt
