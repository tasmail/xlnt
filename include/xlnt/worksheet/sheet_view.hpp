// Copyright (c) 2014-2017 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
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
#include <xlnt/utils/optional.hpp>
#include <xlnt/worksheet/pane.hpp>
#include <xlnt/worksheet/selection.hpp>

namespace xlnt {

/// <summary>
/// Enumeration of possible types of sheet views
/// </summary>
enum class sheet_view_type
{
    normal,
    page_break_preview,
    page_layout
};

/// <summary>
/// Describes a view of a worksheet.
/// Worksheets can have multiple views which show the data differently.
/// </summary>
class XLNT_API sheet_view
{
public:
    /// <summary>
    /// Sets the ID of this view to new_id.
    /// </summary>
    void id(std::size_t new_id)
    {
        id_ = new_id;
    }

    /// <summary>
    /// Returns the ID of this view.
    /// </summary>
    std::size_t id() const
    {
        return id_;
    }

    /// <summary>
    /// Returns true if this view has a pane defined.
    /// </summary>
    bool has_pane() const
    {
        return pane_.is_set();
    }

    /// <summary>
    /// Returns a reference to this view's pane.
    /// </summary>
    struct pane &pane()
    {
        return pane_.get();
    }

    /// <summary>
    /// Returns a reference to this view's pane.
    /// </summary>
    const struct pane &pane() const
    {
        return pane_.get();
    }

    /// <summary>
    /// Removes the defined pane from this view.
    /// </summary>
    void clear_pane()
    {
        pane_.clear();
    }

    /// <summary>
    /// Sets the pane of this view to new_pane.
    /// </summary>
    void pane(const struct pane &new_pane)
    {
        pane_ = new_pane;
    }

    /// <summary>
    /// Returns true if this view has any selections.
    /// </summary>
    bool has_selections() const
    {
        return !selections_.empty();
    }

    /// <summary>
    /// Adds the given selection to the collection of selections.
    /// </summary>
    void add_selection(const class selection &new_selection)
    {
        selections_.push_back(new_selection);
    }

    /// <summary>
    /// Removes all selections.
    /// </summary>
    void clear_selections()
    {
        selections_.clear();
    }

    /// <summary>
    /// Returns the collection of selections as a vector.
    /// </summary>
    std::vector<xlnt::selection> selections() const
    {
        return selections_;
    }

    /// <summary>
    /// Returns the selection at the given index.
    /// </summary>
    class xlnt::selection &selection(std::size_t index)
    {
        return selections_.at(index);
    }

	/// <summary>
	/// Set sheet view scale
	/// </summary>
	void zoom_scale(uint32_t val)
	{
		zoom_scale_ = val;
	}

	/// <summary>
	/// Return sheet view scale
	/// </summary>
	uint32_t zoom_scale() const
	{
		return zoom_scale_;
	}

	/// <summary>
    /// If show is true, grid lines will be shown for sheets using this view.
    /// </summary>
    void show_grid_lines(bool show)
    {
        show_grid_lines_ = show;
    }

    /// <summary>
    /// Returns true if grid lines will be shown for sheets using this view.
    /// </summary>
    bool show_grid_lines() const
    {
        return show_grid_lines_;
    }

	void right_to_left(bool value)
	{
		right_to_left_ = value;
	}

	bool right_to_left() const
	{
		return right_to_left_;
	}
	
	void show_outline_symbols(bool value)
	{
		show_outline_symbols_ = value;
	}

	bool show_outline_symbols() const
	{
		return show_outline_symbols_;
	}

	void show_formulas(bool value)
	{
		show_formulas_ = value;
	}

	bool show_formulas() const
	{
		return show_formulas_;
	}

	void show_row_col_headers(bool value)
	{
		show_row_col_headers_ = value;
	}

	bool show_row_col_headers() const
	{
		return show_row_col_headers_;
	}

	void show_ruler(bool value)
	{
		show_ruler_ = value;
	}

	bool show_ruler() const
	{
		return show_ruler_;
	}

	void show_white_space(bool value)
	{
		show_white_space_ = value;
	}

	bool show_white_space() const
	{
		return show_white_space_;
	}

	void show_zeros(bool value)
	{
		show_zeros_ = value;
	}

	bool show_zeros() const
	{
		return show_zeros_;
	}

	void tab_selected(bool value)
	{
		tab_selected_ = value;
	}

	bool tab_selected() const
	{
		return tab_selected_;
	}

	void window_protection(bool value)
	{
		window_protection_ = value;
	}

	bool window_protection() const
	{
		return window_protection_;
	}

	/// <summary>
    /// If is_default is true, the default grid color will be used.
    /// </summary>
    void default_grid_color(bool is_default)
    {
        default_grid_color_ = is_default;
    }

    /// <summary>
    /// Returns true if the default grid color will be used.
    /// </summary>
    bool default_grid_color() const
    {
        return default_grid_color_;
    }

    /// <summary>
    /// Sets the type of this view.
    /// </summary>
    void type(sheet_view_type new_type)
    {
        type_ = new_type;
    }

    /// <summary>
    /// Returns the type of this view.
    /// </summary>
    sheet_view_type type() const
    {
        return type_;
    }

	/// <summary>
	/// Returns true if this view has a top_left_cell defined.
	/// </summary>
	bool has_top_left_cell() const
	{
		return top_left_cell_.is_set() || (pane_.is_set() && pane_.get().top_left_cell.is_set());
	}

	/// <summary>
	/// Returns the first top left visible cell.
	/// </summary>
	cell_reference top_left_cell() const
	{
		if (top_left_cell_.is_set())
			return top_left_cell_.get();
		return pane_.get().top_left_cell.get();
	}

	/// <summary>
	/// Set the first top left visible cell.
	/// </summary>
	void top_left_cell(const cell_reference& top_left_cell)
	{
		if (pane_.is_set())
			pane_.get().top_left_cell = top_left_cell;
		else 
			top_left_cell_ = top_left_cell;
	}

    /// <summary>
    /// Returns true if this view is equal to rhs based on its id, grid lines setting,
    /// default grid color, pane, and selections.
    /// </summary>
    bool operator==(const sheet_view &rhs) const
    {
        return id_ == rhs.id_
            && show_grid_lines_ == rhs.show_grid_lines_
            && default_grid_color_ == rhs.default_grid_color_
            && pane_ == rhs.pane_
            && selections_ == rhs.selections_;
    }

private:
    /// <summary>
    /// The id
    /// </summary>
    std::size_t id_ = 0;

	/// <summary>
	/// The sheet view scale
	/// </summary>
	uint32_t zoom_scale_ = 100;

    /// <summary>
    /// Whether or not to show grid lines
    /// </summary>
    bool show_grid_lines_ = true;

    /// <summary>
    /// Whether or not to use the default grid color
    /// </summary>
    bool default_grid_color_ = true;

	/// <summary>
	/// Flag indicating whether the sheet is in 'right to left' display mode. 
	/// When in this mode, Column A is on the far right, Column B ;is one column left of Column A, and so on. 
	/// Also, information in cells is displayed in the Right to Left format.
	/// </summary>
	bool right_to_left_ = false;

	/// <summary>
	/// Flag indicating whether the sheet has outline symbols visible. 
	/// This flag shall always override SheetPr element's outlinePr child element whose attribute is named showOutlineSymbols when there is a conflict.
	/// </summary>
	bool show_outline_symbols_ = false;

	/// <summary>
	/// Flag indicating whether this sheet should display formulas.
	/// </summary>
	bool show_formulas_ = false;

	/// <summary>
	/// Flag indicating whether the sheet should display row and column headings.
	/// </summary>
	bool show_row_col_headers_ = true;

	/// <summary>
	/// Show the ruler in Page Layout View.
	/// </summary>
	bool show_ruler_ = false;

	/// <summary>
	/// Flag indicating whether page layout view shall display margins. 
	/// False means do not display left, right, top (header), and bottom (footer) margins (even when there is data in the header or footer).
	/// </summary>
	bool show_white_space_ = false;

	/// <summary>
	/// Flag indicating whether the window should show 0 (zero) in cells containing zero value.
	/// When false, cells with zero value appear blank instead of showing the number zero.
	/// </summary>
	bool show_zeros_ = true;

	/// <summary>
	/// Flag indicating whether this sheet is selected. 
	/// When only 1 sheet is selected and active, this value should be in synch with the activeTab value. 
	/// In case of a conflict, the Start Part setting wins and sets the active sheet tab.
	///
	/// Multiple sheets can be selected, but only one sheet shall be active at one time.
	/// </summary>
	bool tab_selected_ = false;

	/// <summary>
	/// Flag indicating whether the panes in the window are locked due to workbook protection. 
	/// This is an option when the workbook structure is protected.
	/// </summary>
	bool window_protection_ = false;

    /// <summary>
    /// The type of this view
    /// </summary>
    sheet_view_type type_ = sheet_view_type::normal;

    /// <summary>
    /// The optional pane
    /// </summary>
    optional<xlnt::pane> pane_;

    /// <summary>
    /// The collection of selections
    /// </summary>
    std::vector<xlnt::selection> selections_;

	/// <summary>
	/// The first top left visible cell.
	/// </summary>
	optional<cell_reference> top_left_cell_;
};

} // namespace xlnt
