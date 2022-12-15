# Greta Armbrust
# Final Project

# Loops, File IO, Working with external library

def screenplay_final_project(): #TODO I think you want this to be a class.
    """
    sig: NoneType -> NoneType
    """
    
    import os #TODO move imports to the top of the script

    import docx

    from docx import Document

    error_message = 'Error. Not a valid submission. Please run program again.'
    thanks = 'Thank you.'

    print ('Hello, Thank you for using this program.')
    print ('Please select what you would like me to do.')
    print ('A. Ask for formatting for specific screenplay elements')
    print ('B. Format prewritten text')
    print ('C. Make a title page')
    choice = input ('Type the letter of your choice here: ')
    print()

    if choice == 'a' or choice == 'A' : # see if there is a way to allow it to repeat until a user chooses to quit
        element = input ('What element would you like formatting info for? ')
        print()
        
        all_caps = 'It will be in all caps.'
        sentence_case = 'Printed like standard text.'
        no_indents = 'There are no indents used.'
        shooting_script = ('This generally only appears in shooting scripts')
        
        if element == 'Scene Heading':
            print(no_indents)
            print(all_caps)
            print('Example:')
            print('EXT. LIBRARY - DAY')
            print('This means a scene is taking place outside of a library during the day.')
        elif element == 'Action':
            print('Description of a scene which is written in present tense.')
            print(no_indents)
            print(sentence_case)
        elif element == 'Character':
            print('Name of the character placed above dialogue')
            print('Indent 2" on the left')
            print(all_caps)
        elif element == 'Dialogue':
            print('Indent 1" on the left and 1.5" on the right')
            print(sentence_case)
        elif element == 'Parenthetical':
            print('Placed between charater name and dialogue')
            print('Indent 1.5" on the left and 2" on the right')
            print (sentence_case)
        elif element == 'Extension':
            print('Placed after character name in parentheses to tell how a voice is heard')
            print('Example:')
            print('FRANK (V.O.)')
            print ("This means that Frank's voice is heard in voice-over")
        elif element == 'Transition':
            print('These are editing directions')
            print(shooting_script)
            print('Indent 4" on the left')
            print(all_caps)
        elif element == 'Shot':
            print('These specify particular shots')
            print(shooting_script)
            print(no_indents)
            print(all_caps)
        else:
            print (error_message)

    elif choice == 'b' or choice == 'B' :
        print('Before we run this please double check that your text is properly coded.')
        ans = input('Would you like to review the formatting codes? ')
        print()
        
        if ans == 'yes' or ans == 'Yes' or ans == 'Y' or ans == 'y':
            
            folder_location = input('Please input the location of this program on your computer. ')

            print()
            
            format_codes = open(folder_location + '/Formatting_Codes.txt','r')
            for line in format_codes:
                print(line)
            format_codes.close

            def screenplay_format():
                """TODO move this out and up then you can call it from within
                your screenplay_final_project()"""
                """sig: NoneType -> NoneType"""

                from docx import Document ##TODO only need to import once.
                from docx.shared import Inches ##TODO these can be imported together with commas "Inches, Pt"
                from docx.shared import Pt

                print('We will now format your text.')
                start_text = input('Please input the location of this text. ')
                print(thanks)

                start_text = open(start_text, 'r')
                lines = start_text.readlines()
                for line in lines:
                    label, text = line.split(':', 1)
                    text = text.strip()
                      
                ans = input('Should the screenplay be saved as a new document of appended to an old one? ')
                print()
                
                if ans == 'new' or ans == 'New':
                    title = input('What is the title of the screenplay? ')

                    document = Document()
                    document.save(title + '.docx')

                    sections = document.sections
                    for section in sections:
                        section.top_margin = Inches(1.0)
                        section.bottom_margin = Inches(1.0)
                        section.left_margin = Inches(1.5)
                        section.right_margin = Inches(1.0)

                    def line_by_line():
                        """This can be moved up and out as well. """
                        
                        if label=='HEAD': ##TODO I would look at using case instead of the ifs
                            
                            text = text.upper()
                            
                            paragraph = document.add_paragraph(text) ##TODO this can be moved to above the if statements
                            run = document.add_paragraph().add_run()
                            font = run.font
                            font.name = 'Courier'
                            font.size = Pt(12)

                            paragraph_format = paragraph.paragraph_format ##TODO then this can be placed after the if statements

                        if label=='ACTION':

                            paragraph = document.add_paragraph(text)
                            run = document.add_paragraph().add_run()
                            font = run.font
                            font.name = 'Courier'
                            font.size = Pt(12)

                        if label=='CHAR':

                            text = text.upper()

                            paragraph = document.add_paragraph(text)
                            run = document.add_paragraph().add_run()
                            font = run.font
                            font.name = 'Courier'
                            font.size = Pt(12)

                            paragraph_format = paragraph.paragraph_format
                            paragraph_format.left_indent = Inches(2.0)
                            
                        if label=='DIA':
                            
                            paragraph = document.add_paragraph(text)
                            run = document.add_paragraph().add_run()
                            font = run.font
                            font.name = 'Courier'
                            font.size = Pt(12)
                            
                            paragraph_format = paragraph.paragraph_format
                            paragraph_format.left_indent = Inches(1.0)
                            paragraph_format.right_indent = Inches(1.5)

                        if label=='PAREN':

                            paragraph = document.add_paragraph('(' + text + ')')
                            run = document.add_paragraph().add_run()
                            font = run.font
                            font.name = 'Courier'
                            font.size = Pt(12)
                            
                            paragraph_format = paragraph.paragraph_format
                            paragraph_format.left_indent = Inches(1.5)
                            paragraph_format.right_indent = Inches(2.0)

                        if label=='EXTEN': #needs help
                        
                            paragraph = document.add_paragraph(text)
                            run = document.add_paragraph().add_run()
                            font = run.font
                            font.name = 'Courier'
                            font.size = Pt(12)

                        if label=='TRAN':

                            text = text.upper()

                            paragraph = document.add_paragraph(text +':')
                            run = document.add_paragraph().add_run()
                            font = run.font
                            font.name = 'Courier'
                            font.size = Pt(12)
                            
                            paragraph_format = paragraph.paragraph_format
                            paragraph_format.left_indent = Inches(4.0)

                        if label=='SHOT':

                            text = text.upper()

                            paragraph = document.add_paragraph(text +'-')
                            run = document.add_paragraph().add_run()
                            font = run.font
                            font.name = 'Courier'
                            font.size = Pt(12)

                    line_by_line()

                elif ans == 'append' or ans == 'Append':
                    print('Please make sure the file you wish to append is whithin this folder')

                    pre_existing_file = input ("Please input the pre-existing screenplay's filename")

                    document = Document(pre_existing_file)
                    document.save(pre_existing_file)

                    line_by_line()

                else:
                    print (error_message)
                
            screenplay_format()
            
        elif ans == 'no' or ans == 'No' or ans == 'N' or ans == 'n':
            
            screenplay_format()
            
        else:
            print (error_message)

    elif choice == 'c' or choice == 'C':
        print ('We will be formating a title page')

        from docx import Document
        from docx.shared import Inches
        from docx.shared import Pt

        document = Document()

        title = input('What is the title of the screenplay this page is for? ')

        document.save(title + 'Title Page.docx')

        run = document.add_paragraph().add_run()
        font = run.font

        font.name = 'Courier'
        font.size = Pt(12)

        sections = document.sections
            for section in sections:
                section.top_margin = Inches(1.0)
                section.bottom_margin = Inches(1.0)
                section.left_margin = Inches(1.5)
                section.right_margin = Inches(1.0)
        

        # Courier Font
        # Margins left 1.5 right 1.0 top & bottom 1.0
        # Title: All caps centered horizontally, 20-22 lines down (4" from top of page)
        # By-line: 2/4 lines below Title, "by" or "written by"
        # Name: 1 line below

        # Bottom part:
            # company name
            # email
            # phone number


    
        

    else:
        print (error_message)
        
    
        

