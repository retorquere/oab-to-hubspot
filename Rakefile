require 'dotenv/load'
require 'shellwords'
require 'nokogiri'
require 'sqlite3'
require 'json'
require 'vcardio'
require 'rest-client'
require 'hubspot-ruby'

require 'phonelib'
Phonelib.default_country = "NL"

require_relative 'mspack'
require_relative 'parser'

$CACHE = File.expand_path('cache')

task default: 'contacts.json'

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
$LZX = File.join($CACHE, Nokogiri::XML(lzx).remove_namespaces!.at("//Full").text)

file $LZX do
  puts "getting #{$LZX}..."
  Dir[File.join($CACHE, '*.*')].each{|file|
    next if File.basename(file, File.extname(file)) == File.basename($LZX, File.extname($LZX))
    next if File.basename(file) == 'normalized.json'

    puts "- #{file}"
    File.unlink(file)
  }

  system "curl -o #{$LZX.shellescape} -s -u #{user.shellescape} --header 'Content-Type:text/xml' #{(root + '/' + File.basename($LZX)).shellescape}"
  puts $LZX
end

$OAB = File.join($CACHE, File.basename($LZX, File.extname($LZX))) + '.oab'
file $OAB => $LZX do
  puts "getting #{$OAB}..."
  LibMsPack.oab_decompress($LZX, $OAB)
  puts $OAB
end

$JSON = File.join($CACHE, File.basename($OAB, File.extname($OAB))) + '.json'
file $JSON => $OAB do
  puts "getting #{$JSON}..."
  records = OfflineAddressBook::Parser.new($OAB).records.collect{|record|
    record.delete(:Unexepected8D0D)
    record.delete(:Unexepected8C73)
    record[:AddressBookObjectGuid] = record[:AddressBookObjectGuid].inspect if record[:AddressBookObjectGuid] # no idea what's going on here
    record
  }
  open($JSON, 'w'){|f| f.write(JSON.pretty_generate(records)) }
  puts $JSON
end

$NORMALIZED = File.join($CACHE, 'normalized.json')
$_NORMALIZED = File.file?($NORMALIZED) ? JSON.parse(File.read($NORMALIZED)) : {}
at_exit { open($NORMALIZED, 'w'){|f| f.write(JSON.pretty_generate($_NORMALIZED)) } }
def normalize(number)
  if !$_NORMALIZED.has_key?(number)
    phone = Phonelib.parse(number)
    if phone.valid?
      $_NORMALIZED[number] = { 'number' => phone.to_s, 'type' => phone.type }
    else
      $_NORMALIZED[number] = nil
    end
  end

  return $_NORMALIZED[number]
end

file 'contacts.sqlite' => [$JSON, 'Rakefile'] do |t|
  puts "getting #{t.name}..."
  begin
    records = JSON.parse(File.read(t.source))
    File.unlink(t.name) if File.file?(t.name)
    db = SQLite3::Database.new(t.name)
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
    File.unlink(t.name) if File.file?(t.name)
    raise e
  end

  db.execute("DELETE FROM oab WHERE phonetype = 'AssistantTelephoneNumber' AND phonenumber IN (SELECT phonenumber FROM oab WHERE phonetype <> 'AssistantTelephoneNumber')")
  db.execute("DELETE FROM oab WHERE phonetype = 'Business2TelephoneNumber' AND phonenumber IN (SELECT phonenumber FROM oab WHERE phonetype <> 'Business2TelephoneNumber')")
  db.execute("DELETE FROM oab WHERE phonetype <> 'MobileTelephoneNumber' AND phonenumber LIKE '316%' AND phonenumber IN (SELECT phonenumber FROM oab WHERE phonetype = 'MobileTelephoneNumber')")
  puts t.name
end

PHONETYPES = %w{BusinessTelephoneNumber MobileTelephoneNumber Business2TelephoneNumber AssistantTelephoneNumber}
def priority(type)
  raise type unless PHONETYPES.index(type)
  return PHONETYPES.index(type)
end
file 'oab.vcf' => [$JSON, 'Rakefile'] do |t|
  puts "getting #{t.name}"
  oab = JSON.parse(File.read(t.source))

  puts "purging empty and normalizing numbers..."
  contacts = {}
  type = {}
  oab.each{|contact|
    PHONETYPES.each{|field|
      contact[field] = (contact[field] || []).collect{|number| normalize(number) }.compact.uniq
      contact[field].each{|number|
        type[number] = field if !type[number] || priority(field) < priority(type[number])
      }
    }

    numbers = 0
    contact.keys.each{|k|
      if k =~ /TelephoneNumber$/
        numbers += contact[k].length
      elsif contact[k].is_a?(Array) && contact[k].length <= 1
        contact[k] = contact[k].first
      end
    }
    next unless contact['SmtpAddress']
    next unless contact['DisplayName']
    next if numbers == 0
    contact['Surname'] = contact['DisplayName'] if contact['Surname'].nil? && contact['DisplayName'].nil?

    name = [contact['Surname'], contact['GivenName']].compact.join(';')
    if contacts[name]
      PHONETYPES.each{|field|
        contacts[name][field] = (contacts[name][field] + contact[field]).uniq
      }
    else
      contacts[name] ||= contact
    end
  }
  
  open(t.name, 'w'){|f|
    contacts.values.each{|contact|
      numbers = 0
      vcard = VCardio::VCard.new('3.0') do
        fn contact['DisplayName']
        org 'HAN'
        n [contact['Surname'], contact['GivenName']].compact
        email contact['SmtpAddress'], type: 'WORK'
        PHONETYPES.each{|field|
          contact[field].select{|number| type[number] == field }.each{|number|
            numbers += 1
            tel number, type: 'WORK'
          }
        }
      end
      next if numbers == 0

      f.puts(vcard.to_s)
    }
  }
  puts t.name
end

file 'hubspot.json' => [$JSON, 'Rakefile'] do |t|
  puts "getting #{t.name}"
  oab = JSON.parse(File.read(t.source))

  puts "purging empty and normalizing numbers..."
  type = {}
  oab = oab.collect{|contact|
    PHONETYPES.each{|field|
      contact[field] = (contact[field] || []).collect{|number| normalize(number) }.compact.uniq{|number| number['number'] }
      contact[field].each{|number|
        type[number] = field if !type[number] || priority(field) < priority(type[number])
      }
    }

    numbers = 0
    contact.keys.each{|k|
      if k =~ /TelephoneNumber$/
        numbers += contact[k].length
      elsif contact[k].is_a?(Array) && contact[k].length <= 1
        contact[k] = contact[k].first
      end
    }
    if contact['SmtpAddress'].nil?
      nil
    elsif contact['DisplayName'].nil?
      nil
    elsif numbers == 0
      nil
    else
      contact['Surname'] = contact['DisplayName'] if contact['Surname'].nil? && contact['DisplayName'].nil?
      contact['SmtpAddress'].downcase!
      contact
    end
  }.compact

  numbers = []
  oab = oab.collect{|contact|
    keep = false
    PHONETYPES.each{|field|
      contact[field].select{|number| type[number] == field }.each{|number|
        next if numbers.include?(number['number'])

        numbers << number['number']
        keep = true

        if number['type'] == 'mobile' && !contact[:mobilephone]
          contact[:mobilephone] = number['number']
        elsif number['type'] == 'fixed_line' && !contact[:phone]
          contact[:phone] = number['number']
        elsif !contact[:phone]
          contact[:phone] = number['number']
        elsif !contact[:mobilephone]
          contact[:mobilephone] = number['number']
        elsif !contact[:otherphone]
          contact[:otherphone] = number['number']
        else
          raise "No more phone slots for #{contact['DisplayName']}"
        end
      }
    }

    if keep
      {
        firstname: contact['GivenName'],
        lastname: contact['Surname'],
        email: contact['SmtpAddress'],
        mobilephone: contact[:mobilephone],
        phone: contact[:phone],
        # otherphone: contact[:otherphone],
      }
    else
      nil
    end
  }.compact

  open(t.name, 'w'){|f| f.puts(JSON.pretty_generate(oab)) }

  puts 'Updating hubspot...'
  Hubspot.configure(hapikey: ENV['HAPIKEY'])
  oab.each_slice(100).each{|batch|
    print '.'
    Hubspot::Contact.create_or_update!(batch)
  }

  puts t.name
end
