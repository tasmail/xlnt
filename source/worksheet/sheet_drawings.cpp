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
#include <detail/implementations/worksheet_impl.hpp>

namespace xlnt {

	sheet_drawings::sheet_drawings(
		worksheet* parent,
		workbook::drawing_vector& drawings,
		workbook::images_map& images,
		path part
	) :
		parent_(parent),
		drawings_(drawings),
		images_(images),
		part_(part)
	{
	}

	sheet_drawings::sheet_drawings(
		const sheet_drawings& other
	) :
		parent_(other.parent_),
		drawings_(other.drawings_),
		images_(other.images_),
		part_(other.part_)
	{
	}

	sheet_drawings& sheet_drawings::operator=(const sheet_drawings& other)
	{
		this->parent_ = other.parent_;
		this->drawings_ = other.drawings_;
		this->images_ = other.images_;
		this->part_ = other.part_;

		return (*this);
	}

	const workbook::drawing_vector& sheet_drawings::drawings() const
	{
		return drawings_;
	}

	void sheet_drawings::add_drawing(const sheet_drawing& drawing)
	{
		drawings_.push_back(drawing);
		// TODO: get part URI
		// part_ =
		parent_->register_drawings_in_manifest();
	}

	void sheet_drawings::add_drawing_image(
		const sheet_drawing& drawing,
		const workbook::vector_bytes &image,
		const std::string &extension
	)
	{
		sheet_drawing drawing_ = drawing;
		
		// TODO: generate picture content
		// drawing_.picture_path.set(std::string("xl/images/ssss.jpg"));
		add_drawing(drawing_);
	}

	const workbook::vector_bytes &sheet_drawings::get_drawing_image(
		const sheet_drawing& drawing
	) const
	{
		static workbook::vector_bytes empty;
		if (!drawing.picture_rel.is_set())
			return empty;

		const auto relId = drawing.picture_rel.get();

		const auto &manifest = parent_->d_->parent_->manifest();

		if (manifest.has_relationship(part_, relId))
		{
			auto rel = manifest.relationship(part_, relId);
			const auto& it = images_.find(rel.target().path().string());
			if (images_.end() == it)
				return empty;

			return it->second;
		}

		return empty;
	}

} // namespace xlnt
