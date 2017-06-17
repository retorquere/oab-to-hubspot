#!/usr/bin/env ruby

require 'dotenv/load'
require 'exchange-offline-address-book'
require 'phony'
require 'ostruct'
require 'hubspot-ruby'

require 'autodiscover'
require "autodiscover/debug"

Hubspot.configure(hapikey: ENV["HUBSPOT"])

oab = OfflineAddressBook.new(username: ENV['EWS_USERNAME'], password: ENV['EWS_PASSWORD'], email: ENV['EWS_EMAIL'], cachedir: File.dirname(__FILE__), baseurl: ENV['EWS_BASEURL'])

GAL = {}

def normalize(number)
  number = "#{number}".strip
  number = "00#{number}" if number =~ /^31/
  number.gsub!(/^\+/, '00')
  number.gsub!(/^0([1-9])/) { "0031#{$1}" }
  number.gsub!(/^00/, '+')
  Phony.plausible?(number) ? Phony.format(Phony.normalize(number), format: :international).gsub(' ', '') : nil
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

def assign(contact, number, fallback = true)
  parts = Phony.split(Phony.normalize(number))
  mobile = parts[0,2] == ['31', '6']

  if mobile && !contact[:mobilephone]
    contact[:mobilephone] = number

  elsif !mobile && !contact[:phone]
    contact[:phone] = number

  elsif !contact[:mobilephone]
    contact[:mobilephone] = number

  elsif !contact[:phone]
    contact[:phone] = number

  elsif fallback
    fallback = contact[:email].sub('@', '+assistant@')
    GAL[fallback] ||= {
      lastname: 'Assistant to ' + contact[:lastname],
      email: fallback
    }
    assign(GAL[fallback], number, false)

  else
    raise contact[:lastname]

  end
end

GAL.values.each{|contact|
  contact[:secondary] = contact[:secondary] - PRIMARY
  contact[:primary].each{|n|
    assign(contact, n)
  }
  contact[:secondary].each{|n|
    assign(contact, n)
  }

  contact.delete(:primary)
  contact.delete(:secondary)
}

GAL.values.each_slice(100).each{|contacts|
  Hubspot::Contact.create_or_update!(contacts)
}
