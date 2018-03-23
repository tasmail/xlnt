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
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/sheet_drawing.hpp>

namespace xlnt {

class worksheet;

namespace detail
{
	class xlsx_consumer;
}

/// <summary>
/// </summary>
class XLNT_API sheet_drawings
{
public:
	sheet_drawings(
		worksheet* parent_, 
		workbook::drawing_vector& drawings,
		workbook::images_map& images,
		path part);
	
	sheet_drawings(const sheet_drawings& other);

	sheet_drawings& operator=(const sheet_drawings& other);

	const workbook::drawing_vector& drawings() const;	

	void add_drawing_image(
		const sheet_drawing& drawing,
		const workbook::vector_bytes &image,
		const std::string &extension
	);
	
	const workbook::vector_bytes &get_drawing_image(
		const sheet_drawing& drawing
	) const;
private:
	friend class worksheet;
	friend class sheet_drawing;
	friend class detail::xlsx_consumer;

	void register_relations();

	void add_drawing(
		const sheet_drawing& drawing);

	worksheet* parent_;
	workbook::drawing_vector& drawings_;
	workbook::images_map& images_;
	path part_;
};

} // namespace xlnt
