from azure.devops.connection import Connection
from msrest.authentication import BasicAuthentication
import os
import tkinter
from tkinter.messagebox import showinfo, showerror
import tkinter.tix
from tkinter.constants import *
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import re
import pypandoc

selected_items = []
wiki_content = []
wiki_client = ''
project = ''


def store_cred(url, project, token):
    f = open("stored_cred.txt", "w+")
    f.write(url + "\n" + project + '\n' + token)
    f.close()


def print_wiki_page(_path, _wiki_content, project_name):
    global wiki_client
    global project
    wikipage = wiki_client.get_page(project=project_name, path=_path, wiki_identifier='SS365 ADO Wiki Templates',
                                    recursion_level='full', version_descriptor=None)
    modified_string = wikipage.page.path[1:]
    modified_string = modified_string.replace("/", ".")
    _wiki_content.append(modified_string)
    for subpage in wikipage.page.sub_pages:
        _wiki_content = print_wiki_page(subpage.path, _wiki_content, project_name)
    return _wiki_content


def fetch_wiki(mainroot, root, ent1, ent2, ent3, n):
    store_cred(ent1, ent2, ent3)
    # Fill in with your personal access token and org URL
    personal_access_token = ent3
    organization_url = ent1

    # Create a connection to the org
    credentials = BasicAuthentication('', personal_access_token)
    connection = Connection(base_url=organization_url, creds=credentials)

    # Get a client (the "core" client provides access to projects, teams, etc)
    core_client = connection.clients.get_core_client()
    global wiki_client
    wiki_client = connection.clients.get_wiki_client()
    project_name = ent2
    global project
    project = ent2

    global wiki_content
    #try:
    wiki_content = print_wiki_page('/', wiki_content, project_name)
    #except:
        #showerror("Error", "There was an issue fetching the wiki. Please check entered credentials.")
    wiki_content.sort(key=len)
    # print(wiki_content)

    view = View(root)
    root.update()
    b1 = tkinter.Button(root, text="Generate", command=lambda: generate(mainroot, root, view, n), bg='white',
                        relief=tkinter.FLAT,
                        height=2, width=35)
    b1.grid(row=8, column=0, columnspan=2, pady=10)
    root.update()


class View(object):
    def __init__(self, root):
        self.root = root
        self.top = ''
        self.makeCheckList()

    def makeCheckList(self):
        self.top = tkinter.tix.Frame(self.root, relief=tkinter.FLAT, bd=1, bg='white')

        self.top.cl = tkinter.tix.CheckList(self.root, browsecmd=self.selectItem, width=400, height=300)
        self.top.grid(row=6, column=0, columnspan=2, sticky=EW)

        self.top.cl.grid(row=7, column=0, columnspan=2, sticky=EW)
        global wiki_content
        for page in wiki_content:
            if page != '':
                try:
                    self.top.cl.hlist.add(page, text=page.split('.')[-1])
                    self.top.cl.setstatus(page, "off")
                    self.top.cl.hlist.config(fg='grey', bg='white', pady='4', font='segoe')
                except:
                    print("some issue")
        self.top.cl.autosetmode()

    def selectItem(self, item):
        global selected_items
        if self.top.cl.getstatus(item) == 'on':
            if item not in selected_items:
                selected_items.append(item)
        else:
            if item in selected_items:
                selected_items.remove(item)


def generate(mainroot, root, view, n):
    global selected_items
    print(selected_items)
    # selected_items.sort()
    # print(view.top)
    global wiki_client
    global project
    # try:
    b = ''
    for item in selected_items:
        print(item)
        new_item = '/' + item.replace(".", "/")
        wikipage = wiki_client.get_page(project=project, path=new_item, wiki_identifier='SS365 ADO Wiki Templates',
                                        recursion_level='full', version_descriptor=None, include_content=True)
        b = b + wikipage.page.content

    f = open("imported1.md", "w")
    print(repr(b))
    rb = b.replace("[image.png](/.", "")
    rb = rb.replace("[[_TOC_]]", "")
    rb = rb.replace("\t", "")
    rb = rb.replace(r'\n', "")
    rb = rb.replace("<", "")
    rb = rb.replace(">", "")
    rb = rb.replace("]", "")
    rb = rb.replace("[", "")
    rb = rb.replace("!", "")
    rb = re.sub(r'https?:\/\/.*[\r\n]*', '', rb, flags=re.MULTILINE)
    print(repr(rb))
    f.write(rb)
    f.close()
    directory = filedialog.asksaveasfilename(parent=root, defaultextension='.docx')
    print(directory)
    cmd = 'cmd /c "python Test.py ' + directory + ' --template '+ 'doc_templates/'+n+' --files imported1.md"'
    #cmd = 'pandoc -f gfm -t docx imported1.md --reference-doc doc_templates/'+n+' -o wip.docx'
    print(cmd)
    os.system(cmd)
    '''from docxcompose.composer import Composer
    from docx import Document
    master = Document('doc_templates/'+n)
    composer = Composer(master)
    doc1 = Document("wip.docx")
    composer.append(doc1)
    composer.save(directory)'''
    showinfo("Success", "File has been generated!")
    #mainroot.destroy()
    '''except:
        showinfo("Error", "There was an issue generating the file")
        root.destroy()'''


def upload_to_wiki(url, project, token, filename, wikiname):
    from azure.devops.connection import Connection
    from azure.devops.v5_1.wiki import WikiPageCreateOrUpdateParameters
    from msrest.authentication import BasicAuthentication
    # Fill in with your personal access token and org URL
    personal_access_token = token
    # organization_url = 'https://dev.azure.com/YOURORG'
    organization_url = url

    # Create a connection to the org
    credentials = BasicAuthentication('', personal_access_token)
    connection = Connection(base_url=organization_url, creds=credentials)

    # Get a client (the "core" client provides access to projects, teams, etc)
    core_client = connection.clients.get_core_client()
    '''cmd = 'pandoc -f docx -t gfm ' + filename + '.docx -o foo.markdown'
    os.system(cmd)'''
    pypandoc.convert_file(filename + '.docx', to='markdown_github', outputfile="foo.markdown")
    data = ''
    with open('foo.markdown', 'r') as file:
        data = file.read()

    client = connection.clients.get_wiki_client()
    project_name = project
    wikiPageCreateOrUpdateParameters = WikiPageCreateOrUpdateParameters()
    wikiPageCreateOrUpdateParameters.content = data
    a = client.create_or_update_page(project=project, path='/' + wikiname,
                                     wiki_identifier='SS365 ADO Wiki Templates',
                                     version=None, parameters=wikiPageCreateOrUpdateParameters)
    showinfo("Success", "File successfully pushed to wiki!")


def main():
    import os
    file_list = os.listdir(os.path.dirname(__file__) + "\doc_templates")
    print(file_list)
    root = tkinter.tix.Tk()

    root.wm_iconbitmap('logo.ico')
    root.title('DocGen')
    root.geometry("850x600")

    # style.theme_use("MyStyle")
    # Notebook Style
    noteStyler = tkinter.ttk.Style()

    noteStyler.element_create('Plain.Notebook', "from", 'default')
    noteStyler.configure("TNotebook.Tab", padding=[10, 135], relief=tkinter.FLAT, background="darkred")
    noteStyler.map("TNotebook.Tab", background=[("selected", "darkred")], foreground=[("selected", "blue")]);
    noteStyler.configure("TNotebook", tabposition='wn')
    '''COLOR_1 = 'black'
    COLOR_2 = 'white'
    COLOR_3 = 'red'
    COLOR_4 = '#2E2E2E'
    COLOR_5 = '#8A4B08'
    COLOR_6 = '#DF7401'
    noteStyler.element_create('Plain.Notebook', "from", 'default')
    noteStyler.configure("TNotebook", background=COLOR_1, borderwidth=0)
    noteStyler.configure("TNotebook.Tab", background="black", foreground=COLOR_3,
                         lightcolor=COLOR_6, borderwidth=2)
    noteStyler.configure("TFrame", background=COLOR_1, foreground=COLOR_2, borderwidth=0)'''
    tabControl = tkinter.ttk.Notebook(root)
    tab1 = tkinter.ttk.Frame(tabControl, width=200, height=200)
    tab2 = tkinter.ttk.Frame(tabControl, width=200, height=200)
    tabControl.add(tab1, text='Wiki to Document')
    tabControl.add(tab2, text='Document to Wiki')

    tabControl.pack(expand=1, fill="both")
    lab = tkinter.Label(tab1, text="ADO URL")
    lab.grid(row=1, column=0, sticky=tkinter.W)
    # lab.pack()
    ent1 = tkinter.Entry(tab1, width=100, relief=tkinter.FLAT)
    ent1.grid(row=1, column=1, pady=10, padx=5)
    lab2 = tkinter.Label(tab1, text="Project Name", anchor=tkinter.W)
    lab2.grid(row=2, column=0, sticky=tkinter.W)
    ent2 = tkinter.Entry(tab1, width=100, relief=tkinter.FLAT)
    ent2.grid(row=2, column=1, pady=10, padx=5)
    lab3 = tkinter.Label(tab1, text="Access token", anchor=tkinter.W)
    lab3.grid(row=3, column=0, sticky=tkinter.W)
    ent3 = tkinter.Entry(tab1, width=100, relief=tkinter.FLAT)
    ent3.grid(row=3, column=1, pady=10, padx=5)
    lab4 = tkinter.Label(tab1, text="Template", anchor=tkinter.W)
    lab4.grid(row=4, column=0, sticky=tkinter.W)
    n = StringVar(tab1)
    OPTIONS = file_list
    variable = StringVar(tab1)
    variable.set(OPTIONS[0])
    tkinter.ttk.Style().configure('TCombobox', padding=4,relief='flat',borderwidth = 0,shiftrelief = 'flat')
    ent4 = tkinter.ttk.Combobox(tab1, width = 100, textvariable = n)
    ent4['values'] = tuple(file_list)
    #ent4.config(width=100, relief=tkinter.FLAT, bg="white")
    ent4.grid(row=4, column=1, pady=10, padx=5)
    f = open("stored_cred.txt", "r")
    ent1.insert(END, f.readline().rstrip())
    ent2.insert(END, f.readline().rstrip())
    ent3.insert(END, f.readline().rstrip())
    f.close()
    but = tkinter.Button(tab1, text="Fetch wiki",
                         command=lambda: fetch_wiki(root, tab1, ent1.get(), ent2.get(), ent3.get(), n.get()),
                         bg='white', relief=tkinter.FLAT,
                         height=2, width=35).grid(row=5, column=0, columnspan=2, pady=10)
    # but.pack()

    lab = tkinter.Label(tab2, text="ADO URL")
    lab.grid(row=1, column=0, sticky=tkinter.W)
    entry1 = tkinter.Entry(tab2, width=100, relief=tkinter.FLAT)
    entry1.grid(row=1, column=1, pady=10, padx=5)
    label2 = tkinter.Label(tab2, text="Project Name", anchor=tkinter.W)
    label2.grid(row=2, column=0, sticky=tkinter.W)
    entry2 = tkinter.Entry(tab2, width=100, relief=tkinter.FLAT)
    entry2.grid(row=2, column=1, pady=10, padx=5)
    label3 = tkinter.Label(tab2, text="Access token", anchor=tkinter.W)
    label3.grid(row=3, column=0, sticky=tkinter.W)
    entry3 = tkinter.Entry(tab2, width=100, relief=tkinter.FLAT)
    entry3.grid(row=3, column=1, pady=10, padx=5)
    label4 = tkinter.Label(tab2, text="File name", anchor=tkinter.W)
    label4.grid(row=4, column=0, sticky=tkinter.W)
    entry4 = tkinter.Entry(tab2, width=100, relief=tkinter.FLAT)
    entry4.grid(row=4, column=1, pady=10, padx=5)
    label5 = tkinter.Label(tab2, text="Wiki page name", anchor=tkinter.W)
    label5.grid(row=5, column=0, sticky=tkinter.W)
    entry5 = tkinter.Entry(tab2, width=100, relief=tkinter.FLAT)
    entry5.grid(row=5, column=1, pady=10, padx=5)
    f = open("stored_cred.txt", "r")
    entry1.insert(0, f.readline().rstrip())
    entry2.insert(0, f.readline().rstrip())
    entry3.insert(0, f.readline().rstrip())
    f.close()
    button = tkinter.Button(tab2, text="Push to wiki",
                            command=lambda: upload_to_wiki(entry1.get(), entry2.get(), entry3.get(), entry4.get(),
                                                           entry5.get()),
                            bg='white', relief=tkinter.FLAT,
                            height=2, width=35)
    button.grid(row=6, column=0, columnspan=2, pady=10)

    root.mainloop()


if __name__ == '__main__':
    main()
