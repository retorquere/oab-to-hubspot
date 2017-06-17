#!/usr/bin/env ruby

require 'dotenv/load'
require 'exchange-offline-address-book'
require 'json'
require 'hubspot-ruby'

require 'autodiscover'
#require "autodiscover/debug"

require 'phonelib'
Phonelib.default_country = "NL"

Hubspot.configure(hapikey: ENV["HUBSPOT"])

oab = OfflineAddressBook.new(username: ENV['EWS_USERNAME'], password: ENV['EWS_PASSWORD'], email: ENV['EWS_EMAIL'], cachedir: File.dirname(__FILE__), baseurl: ENV['EWS_BASEURL'])

GAL = {}

def normalize(number)
  return nil unless Phonelib.valid?(number)
  return Phonelib.parse(number)
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
  [:primary, :secondary].each{|cat|
    contact[cat].compact!
    contact[cat].uniq!{|n| n.to_s}
  }
  PRIMARY.concat(contact[:primary].collect{|n| n.to_s})

  next if contact[:primary].empty? && contact[:secondary].empty?

  contact[:lastname] = "#{$2} #{$1}".strip if contact[:lastname] =~ /^(#{record.Surname}[^ ]*) (.*)/
  raise "duplicate email address #{contact[:email]}" if GAL[contact[:email]]
  GAL[contact[:email]] = contact
}

def assigned(contact, type, number)
  return false if contact[type]
  contact[type] = number.to_s
  return true
end

def assign(contact, number, fallback = true)
  case number.type
    when :mobile
      return if assigned(contact, :mobilephone, number)
      return if assigned(contact, :phone, number)
    when :fixed_line
      return if assigned(contact, :phone, number)
      return if assigned(contact, :mobilephone, number)
    else
      raise "#{contact[:lastname]}/#{number}: #{number.type.inspect}"
  end

  if fallback
    fallback = contact[:email].sub('@', '+assistant@')
    GAL[fallback] ||= {
      lastname: 'Assistant to ' + contact[:lastname],
      email: fallback
    }
    return assign(GAL[fallback], number, false)
  end

  raise "#{contact[:lastname]}/#{number}: #{number.type.inspect}"
end

GAL.values.each{|contact|
  contact[:secondary] = contact[:secondary].reject{|n| PRIMARY.include?(n.to_s)}
  contact[:primary].each{|n|
    assign(contact, n)
  }
  contact[:secondary].each{|n|
    assign(contact, n)
  }

  contact.delete(:primary)
  contact.delete(:secondary)
}

open('gal.json', 'w'){|f| f.write(JSON.pretty_generate(GAL)) }

GAL.values.each_slice(100).each{|contacts|
  Hubspot::Contact.create_or_update!(contacts)
}
