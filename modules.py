import win32com.client as win32
import subprocess

path_file = r"c:\Users\ehom\Documents\IdeaProjects\Python\Projects\trackChangesQuote\sample\og_scriptures.docx"


def accept_all(docx_path,tcq_path):
    word = win32.gencache.EnsureDispatch("Word.Application")
    doc = word.Documents.Open(docx_path)
    doc.Activate()
    word.ActiveDocument.Revisions.AcceptAll()
    acceptPath = tcq_path+"\\"+"accepted.docx"
    word.ActiveDocument.SaveAs2(acceptPath, FileFormat=16)
    doc.Close()
    word.Application.Quit()
    acceptPath = acceptPath + ".txlf"
    return acceptPath


def reject_all(docx_path, tcq_path):
    word = win32.gencache.EnsureDispatch("Word.Application")
    doc = word.Documents.Open(docx_path)
    doc.Activate()
    word.ActiveDocument.Revisions.RejectAll()
    rejectPath = tcq_path + "\\" + "rejected.docx"
    word.ActiveDocument.SaveAs2(rejectPath, FileFormat=16)
    doc.Close()
    word.Application.Quit()
    rejectPath = rejectPath + ".txlf"
    return rejectPath


def convert_to_txlf(docx_path, sourceLP):
    subprocess.run([r"convertDOC.cmd",
                    "-l", sourceLP, docx_path])


def segment_and_pseudo(txlf_path):
    subprocess.run([r"segmentTxlf.cmd",
                    txlf_path])
    subprocess.run([r"pseudoTranslateTxlf.cmd",
                    txlf_path])


def create_tm_and_update(tm_path, reject_txlf):
    subprocess.run([r"createLuceneTm.cmd",
                    "-p", "de-de", "-t",
                    "file://"+tm_path])
    subprocess.run([r"cleanuptxlf.cmd",
                    "-l", "en-us",
                    "-p", "de-de",
                    "-t", "file://"+tm_path,
                    reject_txlf])


def analyze_accepted(tm_path, accepted_txlf, log_path):
    subprocess.run([r"analyzetxlf.cmd", "-b",
                    "-e>" "log.csv"
                    "-l", "en-us",
                    "-p", "de-de",
                    "-t", "file://"+tm_path,
                    accepted_txlf])