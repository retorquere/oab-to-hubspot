#!/usr/bin/env ruby

require 'dotenv/load'
require 'exchange-offline-address-book'
require 'phony'
require 'ostruct'
require 'hubspot-ruby'

require 'autodiscover'
require "autodiscover/debug"

Hubspot.configure(hapikey: ENV["HUBSPOT"])

oab = OfflineAddressBook.new(username: ENV['USERNAME'], password: ENV['PASSWORD'], email: ENV['EMAIL'], cachedir: File.dirname(__FILE__))

GAL = {}

def normalize(number)
  number = "#{number}".strip
  number = "00#{number}" if number =~ /^31/
  number.gsub!(/^\+/, '00')
  number.gsub!(/^0([1-9])/) { "0031#{$1}" }
  number.gsub!(/^00/, '+')
  Phony.plausible?(number) ? Phony.format(Phony.normalize(number), format: :international) : nil
end

PRIMARY = []
oab.records.each{|record|
  next unless record.SmtpAddress

  contact = {
    lastname: record.DisplayName,
    email: record.SmtpAddress.downcase,
    primary: [],
    secondary: [],
  }


  [
    :BusinessTelephoneNumber,
    :Business2TelephoneNumber,
    :MobileTelephoneNumber,
    :Assistant,
    :AssistantTelephoneNumber,
  ].each{|field|
    next unless record[field]

    kind = field.to_s =~ /Assistant/ ? :secondary : :primary

    if record[field].is_a?(String)
      contact[kind] << normalize(record[field])
    else
      contact[kind].concat(record[field].collect{|n| normalize(n)})
    end
  }
  contact[:primary].compact!
  contact[:primary].uniq!
  contact[:secondary].compact!
  contact[:secondary].uniq!
  PRIMARY.concat(contact[:primary])

  next if contact[:primary].empty? && contact[:secondary].empty?

  contact[:lastname] = "#{$2} #{$1}".strip if contact[:lastname] =~ /^(#{record.Surname}[^ ]*) (.*)/
  raise "duplicate email address #{contact[:email]}" if GAL[contact[:email]]
  GAL[contact[:email]] = contact
}

GAL.values.each{|contact|
  contact[:secondary] = contact[:secondary] - PRIMARY
  contact[:primary].each{|n|
    parts = Phony.split(Phony.normalize(n))
    if parts[0,2] == ['31', '6'] && !contact[:mobilephone]
      contact[:mobilephone] = n
    elsif parts[0,2] != ['31', '6'] && !contact[:phone]
      contact[:phone] = n
    elsif !contact[:phone]
      contact[:phone] = n
    elsif !contact[:mobilephone]
      contact[:mobilephone] = n
    else
      raise contact[:lastname]
    end
  }
  contact[:secondary].each{|n|
    parts = Phony.split(Phony.normalize(n))
    if parts[0,2] == ['31', '6'] && !contact[:mobilephone]
      contact[:mobilephone] = n
    elsif parts[0,2] != ['31', '6'] && !contact[:phone]
      contact[:phone] = n
    elsif !contact[:secondary_phone]
      contact[:secondary_phone]= n
    else
      raise contact[:lastname]
    end
  }

  contact.delete(:primary)
  contact.delete(:secondary)
}

GAL.values.each_slice(100).each{|contacts|
  Hubspot::Contact.create_or_update!(contacts)
}
