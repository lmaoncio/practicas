require 'spreadsheet'
require 'json'

def get_title(parsed, card_id, title)
  parsed['lists'].each_index do |lists|
    if parsed['lists'][lists]['id'] == card_id
      title = parsed['lists'][lists]['name']
    end  
  end
  title
end

def get_archived(parsed, cards)
  if parsed['cards'][cards]['closed']
    archived = 'true'
  else
    archived = 'false'
  end  
  archived
end

def get_members_list(parsed, cards)
  members_list = []
  parsed['cards'][cards]['idMembers'].each_index do |members|
    parsed['members'].each_index do |names|
      if parsed['members'][names]['id'] == parsed['cards'][cards]['idMembers'][members]
        members_list.push(parsed['members'][names]['username'])
      end
    end
  end
  members_list
end

def get_labels_list(parsed, cards)
  labels_list = []
  parsed['cards'][cards]['labels'].each_index do |labels|
    labels_list.push(parsed['cards'][cards]['labels'][labels]['color'] + "-" + parsed['cards'][cards]['labels'][labels]['name'])
  end
  labels_list
end

def get_attachment_list(parsed, cards)
  attachment_list = []
  parsed['cards'][cards]['attachments'].each_index do |attachments|
    attachment_list.push(parsed['cards'][cards]['attachments'][attachments]['name'] + "-" + parsed['cards'][cards]['attachments'][attachments]['url'])
  end
  attachment_list
end

file = File.open(ARGV.first)

data = file.read

parsed = JSON.parse(data)

count = 1

book = Spreadsheet::Workbook.new

sheet = book.create_worksheet(name: parsed['name'])

sheet.row(0).push('ID_TRELLO', 'TITLE', 'DESCRIPTION', 'SHORT_URL', 'URL','title', 'ARCHIVED', 'MEMBERS', 'LABELS', 'ATTACHED_FILES', 'LAST_ACTIVITY_DATE')

parsed['cards'].each_index do |cards|
  
  card_id = parsed['cards'][cards]['idList']
  title = get_title(parsed, card_id, title)
  archived = get_archived(parsed, cards)
  members_list = get_members_list(parsed, cards)
  labels_list = get_labels_list(parsed, cards)
  attachment_list = get_attachment_list(parsed, cards)

  sheet.row(count).push(parsed['cards'][cards]['id'], 
  parsed['cards'][cards]['name'],
  parsed['cards'][cards]['desc'],
  parsed['cards'][cards]['shortLink'],
  parsed['cards'][cards]['url'],
  title,
  archived,
  members_list.to_s,
  labels_list.to_s,
  attachment_list.to_s,
  parsed['cards'][cards]['dateLastActivity'])

  count += 1
end

book.write parsed['name'] + '.xls'

file.close