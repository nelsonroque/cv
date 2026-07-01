function Header(el)
  if el.level <= 3 then
    table.insert(el.content, 1, pandoc.RawInline("latex", "\\needspace{3\\baselineskip}"))
  end
  return el
end
