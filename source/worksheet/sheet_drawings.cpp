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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/worksheet/sheet_drawing.hpp>
#include <xlnt/worksheet/sheet_drawings.hpp>

namespace xlnt {

	sheet_drawings::sheet_drawings(
		worksheet* parent, 
		workbook::drawing_vector& drawings):
		parent_(parent),
		drawings_(drawings)
	{
	}

	sheet_drawings::sheet_drawings(const sheet_drawings& other):
		parent_(other.parent_),
		drawings_(other.drawings_)
	{
	}

	sheet_drawings& sheet_drawings::operator=(const sheet_drawings& other)
	{
		this->parent_ = other.parent_;
		this->drawings_ = other.drawings_;

		return (*this);
	}

	const workbook::drawing_vector& sheet_drawings::drawings() const
	{
		return drawings_;
	}
	
	void sheet_drawings::add_drawing(const sheet_drawing& drawing)
	{
		drawings_.push_back(drawing);
		parent_->register_drawings_in_manifest();
	}

	void sheet_drawings::add_drawing_image(
		const workbook::vector_bytes &image,
		const sheet_drawing& drawing,
		const std::string &extension
	)
	{
	}

	const workbook::vector_bytes &sheet_drawings::get_drawing_image(
		const sheet_drawing& drawing
	)
	{
		static workbook::vector_bytes empty;
		if (!drawing.picture_path.is_set() || !parent_->workbook().has_image(drawing.picture_path.get()))
			return empty;

		return parent_->workbook().get_image(drawing.picture_path.get());
	}

} // namespace xlnt
