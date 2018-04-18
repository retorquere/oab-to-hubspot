#!/usr/bin/env ruby

require 'dotenv/load'
require 'shellwords'
require 'nokogiri'
require 'sqlite3'
require 'json'

require 'phonelib'
Phonelib.default_country = "NL"

require_relative 'mspack'
require_relative 'parser'

def get_redirect(url)
  redirect = `curl -s -I #{url.shellescape}`

  redirect.gsub("\r", '').split("\n").each{|line|
    next unless line =~ /:/
    header, location = *(line.split(':', 2).collect{|v| v.strip })
    return location if header.downcase == 'location'
  }
  raise "No redirect found in #{url}"
end

def pox(email)
  return Nokogiri::XML::Builder.new { |xml|
    xml.Autodiscover('xmlns' => 'http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006') {
      xml.Request {
        xml.EMailAddress(email)
        xml.AcceptableResponseSchema('http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a')
      }
    }
  }.to_xml.gsub("\n", '')
end

redirect = get_redirect('http://autodiscover.han.nl/autodiscover/autodiscover.xml')

user = ENV['EWS_USERNAME'] + ':' + ENV['EWS_PASSWORD']
redirect = `curl -s -d #{pox(ENV['EWS_EMAIL']).shellescape} -u #{user.shellescape} --header 'Content-Type:text/xml' #{redirect.shellescape}`
email = Nokogiri::XML(redirect).remove_namespaces!.at('//RedirectAddr').text

redirect = get_redirect("http://autodiscover.#{email.sub(/.*@/, '')}/autodiscover/autodiscover.xml")

user = ENV['EWS_USERNAME'] + ':' + ENV['EWS_APP_PASSWORD']
root = `curl -s -d #{pox(email).shellescape} -u #{user.shellescape} --header 'Content-Type:text/xml' #{redirect.shellescape}`
root = Nokogiri::XML(root).remove_namespaces!.at("//Protocol[Type[text()='EXCH']]/OABUrl").text
puts root

lzx = `curl -s -u #{user.shellescape} --header 'Content-Type:text/xml' #{(root + '/oab.xml').shellescape}`
lzx = Nokogiri::XML(lzx).remove_namespaces!.at("//Full").text

system "curl -O -s -u #{user.shellescape} --header 'Content-Type:text/xml' #{(root + '/' + lzx).shellescape}" unless File.file?(lzx)

oab = File.basename(lzx, File.extname(lzx)) + '.oab'
LibMsPack.oab_decompress(lzx, oab) unless File.file?(oab)

json = File.basename(lzx, File.extname(lzx)) + '.json'
if !File.file?(json)
  records = OfflineAddressBook::Parser.new(oab).records.collect{|record|
    record.delete(:Unexepected8D0D)
    record.delete(:Unexepected8C73)
    record[:AddressBookObjectGuid] = record[:AddressBookObjectGuid].inspect if record[:AddressBookObjectGuid] # no idea what's going on here
    record
  }
  open(json, 'w'){|f| f.write(JSON.pretty_generate(records)) }
end

def normalize(number)
  return nil unless Phonelib.valid?(number)
  return Phonelib.parse(number).to_s
end

sqlite = File.basename(lzx, File.extname(lzx)) + '.sqlite'
if !File.file?(sqlite)
  begin
    records = JSON.parse(File.read(json))
    db = SQLite3::Database.new(sqlite)
    db.execute('CREATE TABLE oab (displayname, phonetype, phonenumber)')
    records.each_slice(100){|batch|
      db.transaction
      batch.each{|record|
        record.each_pair{|k, v|
          next unless k =~ /TelephoneNumber$/
          v.each{|number|
            number = normalize(number)
            next unless number
            db.execute('INSERT INTO oab (displayname, phonetype, phonenumber) VALUES (?, ?, ?)', record['DisplayName'] || "HAN: #{number}", k, number)
          }
        }
      }
      db.commit
    }
  rescue => e
    File.unlink(sqlite) if File.file?(sqlite)
    raise e
  end
end
db.execute("DELETE from oab where phonetype = 'AssistantTelephoneNumber' and phonenumber in (select phonenumber from oab where phonetype <> 'AssistantTelephoneNumber')")
db.execute("DELETE from oab where phonetype = 'Business2TelephoneNumber' and phonenumber in (select phonenumber from oab where phonetype <> 'Business2TelephoneNumber')")
