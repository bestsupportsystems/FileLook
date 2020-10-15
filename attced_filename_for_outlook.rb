require "byebug"
require "win32ole"

puts Dir.pwd


def each_mail
  ol = WIN32OLE::connect("Outlook.Application")
  myNameSpace = ol.getNameSpace("MAPI")
  folder = myNameSpace.GetDefaultFolder(6)
  byebug
  folder.Items.each.reverse_each do |mail|
    GC.start
    yield mail
  end
end
@cnt = 0

each_mail do |mail|
  @cnt = @cnt + 1
    if mail.Attachments.Count != 0 then
    mail.Attachments.each do |item|
      if URI.unescape(item.FileName).size < item.FileName.size
        puts "メールタイトル:#{mail.Subject}"
        puts "添付ファイル数:#{mail.Attachments.Count}"
        puts "変換前:#{item.FileName}"
        puts "変換後:#{URI.unescape(item.FileName)}"
        puts "変換する？（y/n）"
        ans = gets
        if ans.delete("\n") == "y"
          item.SaveAsFile("#{Dir.pwd}/#{URI.unescape(item.FileName)}")
          puts "ドキュメントフォルダに保存しました"
          puts "終わり？(y/n)"
          ans = gets
          if ans.delete("\n") == "y"
            puts "終わります"
            exit
          end
        end
      end
    end
  end
  exit if @cnt == 100
end
puts "終わります"
