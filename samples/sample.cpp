// Copyright (c) 2017 Thomas Fussell
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

#include <helpers/path_helper.hpp>
#include <xlnt/xlnt.hpp>

void load_image()
{
	xlnt::workbook wb_image;
	wb_image.load(path_helper::sample_file("image.xlsx"));
	auto sheet_drawings = wb_image.active_sheet().sheet_drawings();
	for (const auto &sheet_drawing : sheet_drawings.drawings())
	{
		const auto& data = sheet_drawings.get_drawing_image(sheet_drawing);
		auto picture_name = sheet_drawing.get_picture_name();
		if (picture_name.is_set()) {
			printf(picture_name.get().c_str());
			printf("\n");
			printf("Data lenght: %d", (int)data.size());
			printf("\n");
		}
	}
}

void save_image()
{
	xlnt::workbook wb_image;

	auto sheet_drawings = wb_image.active_sheet().sheet_drawings();

	xlnt::sheet_drawing drawing;
	drawing.from.column_index(1);
	drawing.from.row(3);
	drawing.from_col_offset = 371475;
	drawing.from_row_offset = 76200;

	xlnt::cell_reference to;	
	to.column_index(2);
	to.row(5);
	drawing.to.set(to);
	drawing.to_col_offset.set(75384);
	drawing.to_row_offset.set(8709);

	auto image_path = path_helper::sample_file("cafe.jpg");
	std::ifstream file(image_path.string(), std::ios::binary);
	file.unsetf(std::ios::skipws);

	xlnt::workbook::vector_bytes image_data;
	
	image_data.insert(image_data.begin(),
		std::istream_iterator<std::uint8_t>(file),
		std::istream_iterator<std::uint8_t>());

	sheet_drawings.add_drawing_image(
		drawing,
		image_data,
		"jpg");


	xlnt::sheet_drawing drawing2;
	drawing2.from.column_index(15);
	drawing2.from.row(15);

	to.column_index(21);
	to.row(21);
	drawing2.to.set(to);

	sheet_drawings.add_drawing_image(
		drawing2,
		image_data,
		"jpg");

	wb_image.save(path_helper::sample_file("image~.xlsx"));
}

void sample2()
{
	xlnt::workbook wb;
	xlnt::worksheet ws = wb.active_sheet();

	ws.cell("A1").value(5);
	ws.cell("B2").value("string data");
	ws.cell("C3").formula("=RAND()");

	ws.merge_cells("C3:C4");
	ws.freeze_panes("B2");

	wb.save("sample.xlsx");
}


int main()
{
	save_image();

	load_image();

	sample2();

	return 0;
}
