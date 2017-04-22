// Copyright (C) 2016-2017 Thomas Fussell
// Copyright (C) 2002-2007 Ariya Hidayat (ariya@kde.org).
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions
// are met:
//
// 1. Redistributions of source code must retain the above copyright
// notice, this list of conditions and the following disclaimer.
// 2. Redistributions in binary form must reproduce the above copyright
// notice, this list of conditions and the following disclaimer in the
// documentation and/or other materials provided with the distribution.
//
// THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR
// IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
// OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
// IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT,
// INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
// NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
// DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
// THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
// (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
// THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 
#include <array>
#include <algorithm>
#include <cstring>
#include <fstream>
#include <iostream>
#include <list>
#include <string>
#include <vector>

#include <detail/bytes.hpp>
#include <detail/cryptography/compound_document.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

using xlnt::detail::byte_vector;

using directory_id = std::int32_t;
using sector_id = std::int32_t;

class allocation_table
{
public:
    static const sector_id FreeSector = -1;
    static const sector_id EndOfChainSector = -2;
    static const sector_id AllocationTableSector = -3;
    static const sector_id MasterAllocationTableSector = -4;

    allocation_table::allocation_table()
    {
        resize(128);
    }

    std::size_t allocation_table::count() const
    {
        return data_.size();
    }

    void allocation_table::resize(std::size_t newsize)
    {
        data_.resize(newsize, FreeSector);
    }

    void allocation_table::set(std::size_t index, sector_id value)
    {
        if (index >= count()) resize(index + 1);
        data_[index] = value;
    }

    void allocation_table::setChain(std::vector<sector_id> chain)
    {
        if (chain.size())
        {
            for (std::size_t i = 0; i < chain.size() - 1; i++)
            {
                set(chain[i], chain[i + 1]);
            }

            set(chain[chain.size() - 1], allocation_table::EndOfChainSector);
        }
    }

    std::vector<sector_id> allocation_table::follow(sector_id start) const
    {
        auto chain = std::vector<sector_id>();
        if (start >= count()) return chain;

        auto p = start;

        auto already_exists = [](const std::vector<sector_id> &chain, sector_id item)
        {
            for (std::size_t i = 0; i < chain.size(); i++)
            {
                if (chain[i] == item) return true;
            }

            return false;
        };

        while (p < count())
        {
            if (p == static_cast<std::size_t>(EndOfChainSector)) break;
            if (p == static_cast<std::size_t>(AllocationTableSector)) break;
            if (p == static_cast<std::size_t>(MasterAllocationTableSector)) break;
            if (already_exists(chain, p)) break;
            chain.push_back(p);
            if (data_[p] >= count()) break;
            p = data_[p];
        }

        return chain;
    }

    void allocation_table::load(const byte_vector &data)
    {
        data_ = data.as_vector_of<sector_id>();
    }

    byte_vector allocation_table::save() const
    {
        return byte_vector::from(data_);
    }

    std::size_t allocation_table::size_in_bytes()
    {
        return count() * 4;
    }

    std::size_t sector_size() const
    {
        return sector_size_;
    }

    void sector_size(std::size_t size)
    {
        sector_size_ = size;
    }

private:
    std::size_t sector_size_ = 4096;
    std::vector<sector_id> data_;
};

class header
{
public:
    static const std::uint64_t magic;

    header()
    {
    }

    bool is_valid() const
    {
        if (threshold_ != 4096) return false;
        if (num_sectors_ == 0) return false;
        if ((num_sectors_ > 109) && (num_sectors_ > (num_master_sectors_ * 127) + 109)) return false;
        if ((num_sectors_ < 109) && (num_master_sectors_ != 0)) return false;
        if (short_sector_size_power_ > sector_size_power_) return false;
        if (sector_size_power_ <= 6) return false;
        if (sector_size_power_ >= 31) return false;

        return true;
    }

    void load(byte_vector &data)
    {
        if (data.size() < 512)
        {
            throw xlnt::exception("bad header");
        }

        *this = data.read<header>();

        if (file_id_ != 0xD0CF11E0A1B11AE1)
        {
            throw xlnt::exception("not ole");
        }

        if (!is_valid() || threshold_ != 4096)
        {
            throw xlnt::exception("bad ole");
        }
    }

    byte_vector save() const
    {
        byte_vector out;
        out.write(*this);

        return out;
    }

    std::size_t sector_size() const
    {
        return std::size_t(1) << sector_size_power_;
    }

    std::size_t short_sector_size() const
    {
        return std::size_t(1) << short_sector_size_power_;
    }

    std::vector<sector_id> sectors() const
    {
        return std::vector<sector_id>(first_master_table.begin(), first_master_table.end());
    }

    std::size_t num_master_sectors() const
    {
        return static_cast<std::size_t>(num_master_sectors_);
    }

    sector_id master_table_start() const
    {
        return master_start_;
    }

    sector_id short_table_start() const
    {
        return short_start_;
    }

    sector_id directory_start() const
    {
        return directory_start_;
    }

    std::size_t threshold() const
    {
        return threshold_;
    }

private:
    std::uint64_t file_id_ = 0xD0CF11E0A1B11AE1;
    std::array<std::uint8_t, 16> ignore1 = {{0}};
    std::uint16_t revision_ = 0x003E;
    std::uint16_t version_ = 0x0003;
    std::uint16_t byte_order_ = 0xFEFF;
    std::uint16_t sector_size_power_ = 9;
    std::uint16_t short_sector_size_power_ = 6;
    std::array<std::uint8_t, 10> ignore2 = {{0}};
    std::uint32_t num_sectors_ = 0;
    sector_id directory_start_ = 0;
    std::array<std::uint8_t, 4> ignore3 = {{0}};
    std::uint32_t threshold_ = 4096;
    sector_id short_start_ = 0;
    std::uint32_t num_short_sectors_ = 0;
    sector_id master_start_ = 0;
    std::uint32_t num_master_sectors_ = 0;
    std::array<sector_id, 109> first_master_table = { allocation_table::FreeSector };
};

struct directory_entry
{
    std::array<char16_t, 32> name = {{0}};
    std::uint16_t name_length = 0;

    enum entry_type
    {
        Empty = 0,
        UserStorage = 1,
        UserStream = 2,
        LockBytes = 3,
        Property = 4,
        RootStorage = 5
    } type;

    enum entry_color
    {
        Red = 0,
        Black = 1
    } color;

    directory_id prev = -1;
    directory_id next = -1;
    directory_id child = -1;

    std::array<std::uint8_t, 16> ignore1;
    std::array<std::uint8_t, 4> ignore2;

    std::uint64_t created = 0;
    std::uint64_t modified = 0;

    sector_id first = 0;

    std::uint32_t size = 0;

    std::array<std::uint8_t, 4> ignore3;

    bool operator==(const directory_entry &rhs) const
    {
        return name == rhs.name
            && name_length == rhs.name_length
            && type == rhs.type
            && color == rhs.color
            && prev == rhs.prev
            && next == rhs.next
            && child == rhs.child
            && created == rhs.created
            && modified == rhs.modified
            && first == rhs.first
            && size == rhs.size;
    }
};

class directory_tree
{
public:
    static const directory_id End = -1;

    static void entry_name(directory_entry &entry, std::u16string name)
    {
        if (name.size() > 31)
        {
            name.resize(31);
        }

        std::copy(name.begin(), name.end(), entry.name.begin());
        entry.name[name.size()] = 0;
        entry.name_length = static_cast<std::uint16_t>((name.size() + 1) * 2);
    }

    directory_tree::directory_tree()
        : entries()
    {
        clear();
    }

    void directory_tree::clear()
    {
        entries = { create_root_entry() };
    }

    std::size_t directory_tree::entry_count()
    {
        return entries.size();
    }

    directory_entry &entry(directory_id index)
    {
        return entries[index];
    }

    const directory_entry &entry(const std::u16string &name) const
    {
        return const_cast<directory_tree &>(*this).entry(name);
    }

    directory_entry &entry(const std::u16string &name, bool create)
    {
        if (!name.length())
        {
            throw xlnt::exception("bad name");
        }

        // quick check for "/" (that's root)
        if (name == u"/") return entry(0);

        // split the names, e.g  "/ObjectPool/_1020961869" will become:
        // "ObjectPool" and "_1020961869"
        auto names = std::vector<std::u16string>();
        auto start = std::size_t(0);
        auto end = std::size_t(0);

        if (name[0] == u'/') start++;

        while (start < name.length())
        {
            end = name.find_first_of('/', start);
            if (end == std::string::npos) end = name.length();
            names.push_back(name.substr(start, end - start));
            start = end + 1;
        }

        // start from root
        auto index = directory_id(0);

        for (auto it = names.begin(); it != names.end(); ++it)
        {
            // find among the children of index
            auto chi = children(index);
            std::ptrdiff_t child = 0;

            for (std::size_t i = 0; i < chi.size(); i++)
            {
                auto ce = entry(chi[i]);

                if (std::u16string(ce.name.data()) == *it)
                {
                    child = static_cast<std::ptrdiff_t>(chi[i]);
                }
            }

            // traverse to the child
            if (child > 0)
            {
                index = static_cast<directory_id>(child);
            }
            else
            {
                // not found among children
                if (!create)
                {
                    throw xlnt::exception("not found");
                }

                // create a new entry
                auto parent = index;
                entries.push_back(directory_entry());
                index = static_cast<directory_id>(entry_count() - 1);
                auto &e = entry(index);
                e.first = 0;
                entry(parent).prev = index;
            }
        }

        return entry(index);
    }

    directory_id parent(std::size_t index)
    {
        // brute-force, basically we iterate for each entries, find its children
        // and check if one of the children is 'index'
        for (auto j = directory_id(0); j < static_cast<directory_id>(entry_count()); j++)
        {
            auto chi = children(j);

            for (std::size_t i = 0; i < chi.size(); i++)
            {
                if (chi[i] == index)
                {
                    return static_cast<directory_id>(j);
                }
            }
        }

        return -1;
    }

    std::u16string directory_tree::path(directory_id index)
    {
        // don't use root name ("Root Entry"), just give "/"
        if (index == 0) return u"/";

        auto current_entry = entry(index);

        auto result = std::u16string(entry(index).name.data());
        result.insert(0, u"/");

        auto current_parent = parent(index);

        while (current_parent > 0)
        {
            current_entry = entry(current_parent);

            result.insert(0, std::u16string(current_entry.name.data()));
            result.insert(0, u"/");

            --current_parent;
            index = current_parent;

            if (current_parent <= 0) break;
        }

        return result;
    }

    std::vector<directory_id> directory_tree::children(directory_id index)
    {
        auto result = std::vector<directory_id>();
        auto &e = entry(index);

        if (e.child < entry_count())
        {
            find_siblings(result, e.child);
        }

        return result;
    }


    void directory_tree::load(byte_vector &data)
    {
        entries.clear();

        auto num_entries = data.size() / sizeof(directory_entry);

        for (auto i = std::size_t(0); i < num_entries; ++i)
        {
            auto e = data.read<directory_entry>();

            if ((e.type != directory_entry::UserStream) 
                && (e.type != directory_entry::UserStorage)
                && (e.type != directory_entry::RootStorage))
            {
                throw xlnt::exception("invalid entry");
            }

            entries.push_back(e);
        }
    }


    byte_vector directory_tree::save() const
    {
        auto result = byte_vector();

        for (auto &entry : entries)
        {
            result.write(entry);
        }

        return result;
    }

    // return space required to save this dirtree
    std::size_t directory_tree::size()
    {
        return entry_count() * sizeof(directory_entry);
    }

    directory_entry create_root_entry() const
    {
        directory_entry root;

        entry_name(root, u"Root Entry");
        root.type = directory_entry::RootStorage;
        root.color = directory_entry::Black;
        root.size = 0;

        return root;
    }

private:
    // helper function: recursively find siblings of index
    void find_siblings(std::vector<directory_id> &result, directory_id index)
    {
        auto e = entry(index);

        // prevent infinite loop
        for (std::size_t i = 0; i < result.size(); i++)
        {
            if (result[i] == index) return;
        }

        // add myself
        result.push_back(index);

        // visit previous sibling, don't go infinitely
        auto prev = e.prev;

        if ((prev > 0) && (prev < entry_count()))
        {
            for (std::size_t i = 0; i < result.size(); i++)
            {
                if (result[i] == prev)
                {
                    prev = 0;
                }
            }

            if (prev)
            {
                find_siblings(result, prev);
            }
        }

        // visit next sibling, don't go infinitely
        auto next = e.next;

        if ((next > 0) && (next < entry_count()))
        {
            for (std::size_t i = 0; i < result.size(); i++)
            {
                if (result[i] == next) next = 0;
            }

            if (next)
            {
                find_siblings(result, next);
            }
        }
    }

    std::vector<directory_entry> entries;
};

} // namespace

namespace xlnt {
namespace detail {

class compound_document_impl
{
public:
    compound_document_impl()
    {
        sector_table_.sector_size(header_.sector_size());
        short_sector_table_.sector_size(header_.short_sector_size());
    }

    byte_vector load_sectors(const std::vector<sector_id> &sectors) const
    {
        auto result = byte_vector();
        const auto sector_size = sector_table_.sector_size();

        for (auto sector : sectors)
        {
            auto position = sector_size * (sector + 1);
            result.append(sectors_.data(), position, sector_size);
        }

        return result;
    }

    byte_vector load_short_sectors(const std::vector<sector_id> &sectors) const
    {
        auto result = byte_vector();
        const auto short_sector_size = short_sector_table_.sector_size();
        const auto sector_size = sector_table_.sector_size();

        for (auto sector : sectors)
        {
            auto position = sector * short_sector_size;
            auto master_allocation_table_index = position / sector_size;

            auto sector_data = load_sectors({ short_container_stream_[master_allocation_table_index] });

            auto offset = position % sector_size;
            result.append(sectors_.data(), offset, short_sector_size);
        }

        return result;
    }

    std::vector<sector_id> load_msat(byte_vector &/*data*/)
    {
        const auto sector_size = header_.sector_size();
        auto master_sectors = header_.sectors();

        if (header_.num_master_sectors() > 109)
        {
            auto current_sector = header_.master_table_start();

            for (auto r = std::size_t(0); r < header_.num_master_sectors(); ++r)
            {
                auto msat = load_sectors({ current_sector });
                auto index = 0;

                while (index < (sector_size - 1) / sizeof(sector_id))
                {
                    master_sectors.push_back(msat.read<sector_id>());
                }

                current_sector = msat.read<sector_id>();
            }
        }

        return master_sectors;
    }

    void load(byte_vector &data)
    {
        header_.load(data);

        const auto sector_size = header_.sector_size();
        const auto short_sector_size = header_.short_sector_size();

        sector_table_.sector_size(sector_size);
        short_sector_table_.sector_size(short_sector_size);

        sector_table_.load(load_sectors(load_msat(data)));
        short_sector_table_.load(load_sectors(sector_table_.follow(header_.short_table_start())));
        auto directory_data = load_sectors(sector_table_.follow(header_.directory_start()));
        directory_.load(directory_data);

        auto first_short_sector = directory_.entry(u"Root Entry", false).first;
        short_container_stream_ = sector_table_.follow(first_short_sector);
    }

    byte_vector save() const
    {
        auto result = byte_vector();

        result.append(header_.save().data());
        result.append(sector_table_.save().data());
        result.append(short_sector_table_.save().data());
        result.append(directory_.save().data());

        return result;
    }

    bool has_stream(const std::u16string &filename) const
    {
        //TODO: do this the right way
        try
        {
            directory_.entry(filename);
        }
        catch (xlnt::exception)
        {
            return false;
        }

        return true;
    }

    void add_stream(
        const std::u16string &name,
        const byte_vector &/*data*/)
    {
        auto entry = directory_.entry(name, !has_stream(name));
    }

    byte_vector stream(const std::u16string &name) const
    {
        if (!has_stream(name))
        {
            throw xlnt::exception("document doesn't contain stream with the given name");
        }

        auto entry = directory_.entry(name);
        byte_vector result;

        if (entry.size < header_.threshold())
        {
            result = load_sectors(short_sector_table_.follow(entry.first));
            result.resize(entry.size);
        }
        else
        {
            result = load_sectors(sector_table_.follow(entry.first));
            result.resize(entry.size);
        }

        return result;
    }

private:
    directory_tree directory_;
    header header_;
    allocation_table sector_table_;
    byte_vector sectors_;
    allocation_table short_sector_table_;
    byte_vector short_sectors_;
    std::vector<sector_id> short_container_stream_;
};

compound_document::compound_document()
    : d_(new compound_document_impl())
{
}

compound_document::~compound_document()
{
}

compound_document_impl &compound_document::impl()
{
    return *d_;
}

compound_document_impl &compound_document::impl() const
{
    return *d_;
}


void compound_document::load(std::vector<std::uint8_t> &data)
{
    byte_vector vec(data);
    return impl().load(vec);
}

std::vector<std::uint8_t> compound_document::save() const
{
    return impl().save().data();
}

bool compound_document::has_stream(const std::u16string &filename) const
{
    return impl().has_stream(filename);
}

void compound_document::add_stream(const std::u16string &name, const std::vector<std::uint8_t> &data)
{
    return impl().add_stream(name, data);
}

std::vector<std::uint8_t> compound_document::stream(const std::u16string &name) const
{
    return impl().stream(name).data();
}

} // namespace detail
} // namespace xlnt
