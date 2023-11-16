require "google_drive"

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
# See this document to learn how to create config.json:
# https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md
session = GoogleDrive::Session.from_config("config.json")

# First worksheet of
# https://docs.google.com/spreadsheet/ccc?key=pz7XtlQC-PYx-jrVMJErTcg
# Or https://docs.google.com/a/someone.com/spreadsheets/d/pz7XtlQC-PYx-jrVMJErTcg/edit?usp=drive_web
ws = session.spreadsheet_by_key("1u1UA5GnFGfkjVuFy_mNKwKFRc8bZklSPDGb1792AB9o").worksheets[0]

def make_method_name(s)
  x = s.gsub(/[- +=,*\/\\(){}\[\]]+/, "")
  x.downcase!
  if x.size == 0 or x[0].match("[0-9]")
    "_" + x
  else
    x
  end
end

class Tabela
  attr :skipping
  def initialize(worksheet, x, y)
    @w = worksheet
    @x = x + 1
    @y = y
    @header = @w.rows[x - 1].drop(y - 1)
    @hh = @w.num_rows - x
    @ww = @w.num_cols - y + 1
    @nheaders = @header.map {|x| make_method_name(x)}
    @skipping = (0..@hh-1).find_all do |i|
      r = @w.rows[i + x]
      r.all?("") or r.any? do |s|
        s.downcase!
        s.include?("total")
      end
    end
  end

  class Red
    def initialize(t, i, ww)
      @t = t
      @i = i
      @ww = ww
    end

    def [] (j)
      @t[@i, j]
    end

    def []=(j, s)
      @t[@i, j] = s.to_s
      @t.save
    end
    
    include Enumerable

    def each()
      if block_given?
        (0 .. @ww-1).each do |j|
          yield @t[@i, j]
        end
      else
        to_enum(:each)
      end
    end
  end

  class Kolona
    def initialize(t, j, hh, skipping)
      @t = t
      @j = j
      @hh = hh
      @skipping = skipping
      @ncells = self.map {|x| make_method_name(x)}
    end

    def [] (i)
      @t[i, @j]
    end

    def []=(i, s)
      @t[i, @j] = s.to_s
      @t.save
    end

    include Enumerable

    def each()
      if block_given?
        (0 .. @hh-@skipping.size-1).each do |i|
          yield @t[i, @j]
        end
      else
        to_enum(:each)
      end
    end

    def method_missing(name, *args)
      i = @ncells.find_index(name.to_s.downcase)
      if i
        if args.size == 0
          @t.row(i)
        else
          raise RuntimeError, "unexpected arguments"
        end
      else
        raise RuntimeError, "row not found"
      end
    end

    def sum()
      self.map {|n| n.to_i}.sum
    end

    def avg()
      self.sum() / (@hh - @skipping.size)
    end
  end

  def [] (i, *args)
    if args.size == 0
      Kolona.new(self, @header.find_index(i), @hh, @skipping)
    elsif args.size == 1
      @skipping.each do |x|
        if x <= i
          i += 1
        else
          break
        end
      end
      @w[i + @x, args[0] + @y]
    else
      raise RuntimeError
    end
  end

  def []=(i, j, s)
    @w[i + @x, j + @y] = s.to_s
    @w.save
  end

  def row(i)
    Red.new(self, i, @ww)
  end

  include Enumerable

  def each()
    if block_given?
      (0 .. @hh-@skipping.size-1).each do |i|
        (0 .. @ww-1).each do |j|
          yield self[i, j]
        end
      end
    else
      to_enum(:each)
    end
  end

  def method_missing(name, *args)
    j = @nheaders.find_index(name.to_s.downcase)
    if j
      if args.size == 0
        Kolona.new(self, j, @hh, @skipping)
      else
        raise RuntimeError, "unexpected arguments"
      end
    else
      raise RuntimeError, "column not found"
    end
  end
end


def main(ws)
  t = Tabela.new(ws, 4, 5)

  p "t", t
  p "t.row(0)", t.row(0)
  p "t.row(0)[0]", t.row(0)[0]
  p "t.row(0).to_a", t.row(0).to_a
  p "t[0, 0]", t[0, 0]
  p "t.map {|x| x + x}", t.map {|x| x + x}
  p "t[\"Prva Kolona\"]", t["Prva Kolona"]
  p "t[\"Prva Kolona\"][0]", t["Prva Kolona"][0]
  p "t[\"Prva Kolona\"].to_a", t["Prva Kolona"].to_a
  p "t.prvaKolona", t.prvaKolona
  p "t.prvaKolona[0]", t.prvaKolona[0]
  p "t.prvaKolona.to_a", t.prvaKolona.to_a
  p "t.prvaKolona.sum", t.prvaKolona.sum
  p "t.prvaKolona.avg", t.prvaKolona.avg
  p "t.prvaKolona._4", t.prvaKolona._4
  p "t.prvaKolona._4.to_a", t.prvaKolona._4.to_a
  p "t.prvaKolona.map { |cell| cell.to_i + 1 }", t.prvaKolona.map { |cell| cell.to_i + 1 }
end

main ws
