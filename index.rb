require 'docx'
require 'mustache'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('admission.docx')

def params 
  {
    student: {
      name: "Misty Abbott",
      instrument: 'xylophone'
    },
    gender_pronoun: 'her',
    gender_pronoun_2: 'she',
    program_period: 'fall 2018 semester',
    school: {
      phone: '(1)2-9323-5342',
      email: ' admissions@schoolofrock.edu'
    },
    event_1: {
      date: 'July, 21st',
      time: '10:00 PM',
      name: 'Admission Form deadline'
    },
    event_2: {
      date: 'Aug, 8th',
      time: '08:00 AM',
      name: 'First Class'
    }
  }
end

def has_only_closed_variables?(text)
  !text.match(/{{[^{}]*}}[^{}]*$/).nil? 
end

def has_open_variable?(text)
  !text.match(/{{[^}]*$/).nil?
end

def render_text(text)
  Mustache.render(text, params)
end

# Retrieve and display paragraphs
doc.paragraphs.each do |p|
  p.text_runs.inject(nil) do |last_tr, current_tr|
    if last_tr.nil?
      if has_open_variable?(current_tr.text)
        current_tr
      else 
        current_tr.text = render_text(current_tr.text) if has_only_closed_variables?(current_tr.text)
        nil
      end
    else
      current_tr_text = current_tr.text
      last_tr.text += current_tr_text
      current_tr.node.remove
      if has_only_closed_variables?(last_tr.text)
        last_tr.text = render_text(last_tr.text)
        nil
      else
        last_tr
      end
    end
  end
end

doc.save('admission-filled.docx')
