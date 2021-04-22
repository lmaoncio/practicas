require 'spreadsheet'
require 'json'

#PARSED DATA
begin
  FILE = File.open(ARGV.first)
  DATA = JSON.parse(FILE.read)
rescue
  puts 'Error: Reading file.'
end

#CONSTANTS
LISTS = 'lists'
ID = 'ID'
NAME = 'name'
ID_LIST = 'idList'
CARDS = 'cards'
CLOSED = 'closed'
ID_MEMBERS = 'idMembers'
MEMBERS = 'members'
USERNAME = 'username'
LABELS = 'labels'
COLOR = 'color'
ATTACHMENTS = 'attachments'
DESCRIPTION = 'desc'
SHORT_LINK = 'shortLink'
URL = 'url'
DATE_LAST_ACTIVITY = 'dateLastActivity'
#PUBLIC VARIABLE
$card_index = 0

#GETS CARD TITLE
def card_title()
  title = ''
   DATA[LISTS].each_index do |item|
    if DATA[LISTS][item][ID] == DATA[CARDS][$card_index][ID_LIST]
      title = DATA[LISTS][item][NAME]  
    end
  end
  title
end

#GETS CARD STATUS
def card_status()
  if DATA[CARDS][$card_index][CLOSED]
    archived = 'true'
  else
    archived = 'false'
  end  
  archived
end

#GETS CARD MEMBER LIST
def get_members_list()
  members_list = []
  DATA[CARDS][$card_index][ID_MEMBERS].each_index do |members|
    DATA[MEMBERS].each_index do |names|
      #puts DATA[MEMBERS][names][ID]
      #puts DATA['members'][names]['id']
      if DATA['members'][names]['id'] == DATA[CARDS][$card_index][ID_MEMBERS][members]
        members_list.push(DATA[MEMBERS][names][USERNAME])
      end
    end
  end
  members_list = members_list.join(",")

  members_list
end

#GETS CARD LABEL LIST
def get_labels_list()
  labels_list = []
  DATA[CARDS][$card_index][LABELS].each_index do |item|
    labels_list.push(DATA[CARDS][$card_index][LABELS][item][COLOR] + "-" + DATA[CARDS][$card_index][LABELS][item][NAME])
  end
  labels_list = labels_list.join(',')
  labels_list
end

#GETS CARD ATTACHMENT LIST
def get_attachment_list()
  attachment_list = []
  DATA[CARDS][$card_index][ATTACHMENTS].each_index do |item|
    attachment_list.push(DATA[CARDS][$card_index][ATTACHMENTS][item][NAME] + "-" + DATA[CARDS][$card_index][ATTACHMENTS][item][URL])
  end
  attachment_list = attachment_list.join(',')
  attachment_list
end

#ROW COUNT
count = 1

#EXCEL CREATION
begin
  book = Spreadsheet::Workbook.new
  sheet = book.create_worksheet(name: DATA[NAME])
  sheet.row(0).push('ID_TRELLO', 'TITLE', 'DESCRIPTION', 'SHORT_URL', 'URL','title', 'ARCHIVED', 'MEMBERS', 'LABELS', 'ATTACHED_FILES', 'LAST_ACTIVITY_DATE')
rescue
  puts 'Error: Creating excel.'
end

#CREATING FORMAT
format = Spreadsheet::Format.new :text_wrap => true

#PARSING TRELLO CARDS
begin
  DATA[CARDS].each_index do |item|
    $card_index = item
    
    #GETTING DATA
    title = card_title()
    archived = card_status()
    members_list = get_members_list()
    labels_list = get_labels_list()
    attachment_list = get_attachment_list()

    #puts DATA[CARDS][item][ID]
    #puts DATA['cards'][item]['id']
    
    #WRITING EXCEL
    row = sheet.row(count).push(DATA['cards'][item]['id'], 
    DATA[CARDS][item][NAME],
    DATA[CARDS][item][DESCRIPTION],
    DATA[CARDS][item][SHORT_LINK],
    DATA[CARDS][item][URL],
    title,
    archived,
    members_list.to_s,
    labels_list.to_s,
    attachment_list.to_s,
    #PARSING DATE
    Date.parse(DATA[CARDS][$card_index][DATE_LAST_ACTIVITY]))

    #FORMATING CELL
    row.set_format(9, format)

    count += 1
  end

  #WRITES XLS AND CLOSES FILE
  book.write DATA[NAME] + '.xls'
rescue
  puts 'Error: Writing data.'
end

FILE.close