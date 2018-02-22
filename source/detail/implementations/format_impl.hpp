#pragma once

#include <cstddef>
#include <sstream>

#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

class alignment;
class border;
class fill;
class font;
class number_format;
class protection;

namespace detail {

struct stylesheet;

struct format_impl
{
	stylesheet *parent;

	std::size_t id;

	optional<std::size_t> alignment_id;
    bool alignment_applied = false;

	optional<std::size_t> border_id;
    bool border_applied = false;

	optional<std::size_t> fill_id;
    bool fill_applied = false;

	optional<std::size_t> font_id;
    bool font_applied = false;

	optional<std::size_t> number_format_id;
    bool number_format_applied = false;

	optional<std::size_t> protection_id;
    bool protection_applied = false;

    bool pivot_button_ = false;
    bool quote_prefix_ = false;

	optional<std::string> style;

    std::size_t references = 0;

	XLNT_API friend bool operator==(const format_impl &left, const format_impl &right)
	{
		return left.parent == right.parent
			&& left.alignment_id == right.alignment_id
			&& left.alignment_applied == right.alignment_applied
			&& left.border_id == right.border_id
			&& left.border_applied == right.border_applied
			&& left.fill_id == right.fill_id
			&& left.fill_applied == right.fill_applied
			&& left.font_id == right.font_id
			&& left.font_applied == right.font_applied
			&& left.number_format_id == right.number_format_id
			&& left.number_format_applied == right.number_format_applied
			&& left.protection_id == right.protection_id
			&& left.protection_applied == right.protection_applied
			&& left.pivot_button_ == right.pivot_button_
			&& left.quote_prefix_ == right.quote_prefix_
			&& left.style == right.style;
	}
	
	XLNT_API friend bool operator<(const format_impl &left, const format_impl &right)
	{
		if (left.parent != right.parent)
			return left.parent < right.parent;
		
		if (left.alignment_id != right.alignment_id)
			return left.alignment_id < right.alignment_id;
		
		if (left.alignment_applied == right.alignment_applied)
			return left.alignment_applied < right.alignment_applied;
		
		if (left.border_id == right.border_id)
			return left.border_id < right.border_id;
		
		if (left.border_applied == right.border_applied)
			return left.border_applied <= right.border_applied;
		
		if (left.fill_id == right.fill_id)
			return left.fill_id < right.fill_id;
		
		if (left.fill_applied == right.fill_applied)
			return left.fill_applied < right.fill_applied;
		
		if (left.font_id == right.font_id)
			return left.font_id < right.font_id;
		
		if (left.font_applied == right.font_applied)
			return left.font_applied < right.font_applied;
		
		if (left.number_format_id == right.number_format_id)
			return left.number_format_id < right.number_format_id;
		
		if (left.number_format_applied == right.number_format_applied)
			return left.number_format_applied < right.number_format_applied;
		
		if (left.protection_id == right.protection_id)
			return left.protection_id < right.protection_id;
		
		if (left.protection_applied == right.protection_applied)
			return left.protection_applied < right.protection_applied;
		
		if (left.pivot_button_ == right.pivot_button_)
			return left.pivot_button_ < right.pivot_button_;
		
		if (left.quote_prefix_ == right.quote_prefix_)
			return left.quote_prefix_ < right.quote_prefix_;
		
		if (left.style == right.style)
			return left.style < right.style;

		return false;
	}

	std::string to_string() const
	{
		std::ostringstream ss;

		ss << "id:'" << id << "' ";

		if (alignment_id.is_set())
			ss << "alignment_id:'" << alignment_id.get() << "' ";
		
		if (border_id.is_set())
			ss << "border_id:'" << border_id.get() << "' ";
		
		if (fill_id.is_set())
			ss << "fill_id:'" << fill_id.get() << "' ";
		
		if (font_id.is_set())
			ss << "font_id:'" << font_id.get() << "' ";
		
		if (number_format_id.is_set())
			ss << "number_format_id:'" << number_format_id.get() << "' ";
		
		if (protection_id.is_set())
			ss << "protection_id:'" << protection_id.get() << "' ";
		
		if (style.is_set())
			ss << "style:'" << style.get() << "' ";

		return ss.str();
	}
};

} // namespace detail
} // namespace xlnt
