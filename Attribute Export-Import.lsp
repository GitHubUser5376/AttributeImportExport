;;; ====================================================================================== ;;;
;;; Scope ========================================================================== Scope ;;;
;
; Seemlessly transfer attributes between an external documents and cad software
;
;;; Scope ========================================================================== Scope ;;;
;;; ====================================================================================== ;;;
;;; Objective List ======================================================== Objective List ;;;
;
; 01) Scope ------------------------- Completed
; 02) Object List ------------------- Completed
; 03) Notes ------------------------- Completed
; 04) Global Variables -------------- Completed
; 05) User Interaction -------------- Completed
; 06) Default Settings -------------- Completed
; 07) Main ATIN --------------------- Completed
; 08) Main ATOUT -------------------- Completed
; 09) Get Attributes ---------------- Completed
; 10) Set/Push Block Attributes ----- Completed
; 11) Set/Push External Document ---- Completed
; 12) Get/Pull External Document ---- Completed
; 13) Other Functions --------------- Completed
; 14) Function Use Global Variables - Completed
; 15) Messages ---------------------- Completed
;
;;; Objective List ======================================================== Objective List ;;;
;;; ====================================================================================== ;;;
;;; Comments ==================================================================== Comments ;;;

; Error and Debugging Help:
; If an unforseen error occurs while using one of the "User Interactions" functions, change
; "(setq *bProgrammerDebug* nil)" to "(setq *bProgrammerDebug* T)" to receive a detailed 
; overview of which functions are being called. The location of the data dump is stored at
; "%LocalAppData%\Temp\AttErrDumpFile.log."

; Created Files:
; A file to remember the any changed the user selects for their preference is created at
; "%AppData%\CadAddins" under the file name "AT-Import-Export-Defaults.txt." Additional log
; files containing errors are stored at "%LocalAppData%\Temp." The name of the log files 
; will either be "AttErrDumpFile.log" or the name of the active document, plus the date and 
; time. 

; Excel Export: 
; Directly eporting and importing information per cell between AutoLisp and Excel was found 
; to be time consuming. As a result, exporting and importing information with Excel is done 
; indirectly with a csv/txt file. This file uses csv by default, but this can be changed to
; txt by going to "C:\Users\[User]\AppData\Roaming\CadAddins\AT-Import-Export-Defaults.txt"
; and change "Excel Type = csv" to "Excel Type = txt."

; Disclaimer:
; All of the code within this file has either been self typed or found as publically 
; available online. To the extent of my knowledge, all sources where code was copied and 
; pasted have a reference either above the function of the code or within the code itself. 

;;; Comments ==================================================================== Comments ;;;
;;; ====================================================================================== ;;;
;;; Global Variables ==================================================== Global Variables ;;;

;; Regular Variables
(vl-load-com)
(setq *ExternalDocument* nil)
(setq *ExcelApp* nil)
(setq *ExternalFileType* nil)
(setq *sFileName* nil)

;; Error Variables
(setq *bProgrammerDebug* nil);<-- Set to T to start debugging error. Set to nil when done
;(setq *error* nil);---------; If there's a custom, default error state, then comment this out
(setq *OriginErrState* nil);-; Stores the original error state when this file is used.
(setq *sErrorMessage* "");---; This is used as a User message when *bProgrammerDebug* is nil, 
;----------------------------; and it provides the process tree during debugging. 
(setq *sDebugMessage* "");---; Provides a glimps of information into the module that failed.
(setq *sErrDumpPath* "");----; Redefined after functions are defined.
(setq *iIterErr* 0);---------; Interates each time the error function is ran.
(setq *iDepthErr* 0);--------; Increments each time a function is ran, and decrements when a 
;----------------------------; function ends

;;; Global Variables ==================================================== Global Variables ;;;
;;; ====================================================================================== ;;;
;;; User Interactions ================================================== User Interactions ;;;

;;; NOTICE: These functions run before *error* is assigned to fcnErrorFunction.

;;; User Interaction - Help Guide
(defun C:ATHELP() 
    (princ "\n \n")
    (princ "\nATHELP - Provides a description of each of the commands added in this lisp file.")
    (princ "\n \n")
    (princ "\nATIN -------- By default ATIN is set to the same input as ATIN-ALL.")
    (princ "\nATIN-All ---- Pulls in all blocks from both the specified external file and active document.")
    (princ "\nATIN-Name --- User types the name of the block to update. Every block with the same name will be updated.")
    (princ "\nATIN-Select - User selects the blocks in the active document to update.")
    (princ "\n \n")
    (princ "\nATOUT -------- By default ATOUT is set to the same input as ATOUT-SELECT.")
    (princ "\nATOUT-All ---- Copies all of the block attributes in the active document and paste them into an external file.")
    (princ "\nATOUT-Name --- Copies the user specified (by name) block attributes in the active document and paste them into an external file.")
    (princ "\nATOUT-Select - Copies the user selected block attributes in the active document and paste them into an external file.")
    (princ "\n \n")
    (princ "\nATPort -------------------- Runs ATPort-With")
    (princ "\nATPort-Release-Apps ------- Resets the status of open applications and releases them from DraftSight's memory.")
    (princ "\nATPort-Release-Apps-Force - Connects applications to DraftSight and then activates \"ATPort-Release-Apps\" command.")
    (princ "\nATPort-Reset-Settings ----- Reset the default states of ATIN, ATOUT, and ATPORT to the starting state.")
    (princ "\nATPort-Set ---------------- Sets the default state of ATIN or ATOUT to one of the extension names listed.")
    (princ "\nATPort-With --------------- Changes the external file output of ATOUT to the file type selected.")
    (princ "\n \n")
    (princ "\nATIN   is currently set to mimic ATIN-")(princ (fcnLoadDefault "C:ATIN"))(princ ".")
    (princ "\nATOUT  is currently set to mimic ATOUT-")(princ (fcnLoadDefault "C:ATOUT"))(princ ".")
    (princ "\nATPort is currently set to work with ")(princ (fcnLoadDefault "C:ATPort-With"))(princ " files.")
    (terpri)(princ)
);ATHELP

;;; User Interaction - Command Inputs
(defun C:ATIN()        
    (fcnResetError)
    (setq *OriginErrState* *error* *error* fcnErrorFunction); Start Error Collection
    (if *bProgrammerDebug* (fcnBuildError "ATIN - Start" 0))
    (fcnMainImport (fcnBlockSelect (fcnLoadDefault "C:ATIN"))); Calling the main function
    (setq *error* *OriginErrState*); End Custom Error
    (setq fcnErrorFunction fcnErrorFunction)
    (if *bProgrammerDebug* (progn ; Display error
        (fcnBuildError "ATIN - End\n***Program Completed***" 0)
        (fcnErrorFunction "")
    ));if<-progn
    (princ)
);defun ATIN

(defun C:ATIN-All()    
    (fcnResetError)
    (setq *OriginErrState* *error* *error* fcnErrorFunction); Start Error Collection
    (if *bProgrammerDebug* (fcnBuildError "ATIN-All - Start" 0))
    (fcnMainImport (fcnBlockSelect "All")); Calling the main function
    (setq *error* *OriginErrState*); End Custom Error
    (setq fcnErrorFunction fcnErrorFunction)
    (if *bProgrammerDebug* (progn ; Display error
        (fcnBuildError "ATIN-All - End\n***Program Completed***" 0)
        (fcnErrorFunction "")
    ));if<-progn
    (princ)
);defun C:ATIN-All

(defun C:ATIN-Name()   
    (fcnResetError)
    (setq fcnErrorFunction fcnErrorFunction)
    (setq *OriginErrState* *error* *error* fcnErrorFunction); Start Error Collection
    (if *bProgrammerDebug* (fcnBuildError "ATIN-Name - Start" 0))
    (fcnMainImport (fcnBlockSelect "Name")); Calling the main function
    (setq *error* *OriginErrState*); End Custom Error
    (setq fcnErrorFunction fcnErrorFunction)
    (if *bProgrammerDebug* (progn ; Display error
        (fcnBuildError "ATIN-Name - End\n***Program Completed***" 0)
        (fcnErrorFunction "")
    ));if<-progn
    (princ)
);defun ATIN-All

(defun C:ATIN-Select() 
    (fcnResetError)
    (setq fcnErrorFunction fcnErrorFunction)
    (setq *OriginErrState* *error* *error* fcnErrorFunction); Start Error Collection
    (if *bProgrammerDebug* (fcnBuildError "ATIN-Select - Start" 0))
    (fcnMainImport (fcnBlockSelect "Select")); Calling the main function
    (setq *error* *OriginErrState*); End Custom Error
    (setq fcnErrorFunction fcnErrorFunction)
    (if *bProgrammerDebug* (progn ; Display error
        (fcnBuildError "ATIN-Select - End\n***Program Completed***" 0)
        (fcnErrorFunction "")
    ));if<-progn
    (princ)
);defun ATIN-Select

(defun C:ATOUT()       
    (fcnResetError)
    (setq *OriginErrState* *error* *error* fcnErrorFunction); Start Error Collection
    (if *bProgrammerDebug* (fcnBuildError "ATOUT - Start" 0))
    (fcnMainExport (fcnBlockSelect (fcnLoadDefault "C:ATOUT"))); Calling the main function
    (setq *error* *OriginErrState*);End Custom Error
    (setq fcnErrorFunction fcnErrorFunction)
    (if *bProgrammerDebug* (progn ; Display error
        (fcnBuildError "ATOUT - End\n***Program Completed***" 0)
        (fcnErrorFunction "")
    ));if<-progn
    (princ)
);defun C:ATOUT

(defun C:ATOUT-All()   
    (fcnResetError)
    (setq *OriginErrState* *error* *error* fcnErrorFunction); Start Error Collection
    (if *bProgrammerDebug* (fcnBuildError "ATOUT-All - Start" 0))
    (fcnMainExport (fcnBlockSelect "All")); Calling the main function
    (setq *error* *OriginErrState*);End Custom Error
    (setq fcnErrorFunction fcnErrorFunction)
    (if *bProgrammerDebug* (progn ; Display error
        (fcnBuildError "ATOUT-All - End\n***Program Completed***" 0)
        (fcnErrorFunction "")
    ));if<-progn
    (princ)
);defun C:ATOUT-All

(defun C:ATOUT-Name()  
    (fcnResetError)
    (setq *OriginErrState* *error* *error* fcnErrorFunction); Start Error Collection
    (if *bProgrammerDebug* (fcnBuildError "ATOUT-Name - Start" 0))
    (fcnMainExport (fcnBlockSelect "Name")); Calling the main function
    (setq *error* *OriginErrState*);End Custom Error
    (setq fcnErrorFunction fcnErrorFunction)
    (if *bProgrammerDebug* (progn ; Display error
        (fcnBuildError "ATOUT-Name - End\n***Program Completed***" 0)
        (fcnErrorFunction "")
    ));if<-progn
    (princ)
);defun C:ATOUT-Name

(defun C:ATOUT-Select()
    (fcnResetError)
    (setq *OriginErrState* *error* *error* fcnErrorFunction); Start Error Collection
    (if *bProgrammerDebug* (fcnBuildError "ATOUT-Select - Start" 0))
    (fcnMainExport (fcnBlockSelect "Select")); Calling the main function
    (setq *error* *OriginErrState*);End Custom Error
    (setq fcnErrorFunction fcnErrorFunction)
    (if *bProgrammerDebug* (progn ; Display error
        (fcnBuildError "ATOUT-Select - End\n***Program Completed***" 0)
        (fcnErrorFunction "")
    ));if<-progn
    (princ)
);defun C:ATOUT-Select

;;; User Interaction - Settings for Importing and exporting
(defun C:ATPort()(C:ATPort-With))
(defun C:ATPort-Release-Apps($bSilent)
    (if (and (not *ExcelApp*)(not *ExternalDocument*)(not $bSilent))
        (princ (strcat "\nNo hanging applications to restore.\n"
            "If an application stops working when ATIN or ATOUT " 
            "is ran, this function will restore them.\n"))
    );if
    (if *ExcelApp* (progn
        (vla-put-visible *ExcelApp* :vlax-true);-----------------; Makes Excel visible
        (vlax-put-property *ExcelApp* 'DisplayAlerts :vlax-true);; Enabling Alerts to be displayed
        (vlax-put-property *ExcelApp* 'EnableEvents :vlax-true);-; Enabling Automatic Events
        (vlax-put-property *ExcelApp* 'ScreenUpdating :vlax-true); Enabling Excel Screen Updates
        (vlax-put-property *ExcelApp* 'Calculation -4105);-------; Switching Calculations to Automatic
        (vlax-release-object *ExcelApp*)(gc)
        (setq *ExcelApp* nil)
        (if (not $bSilent)(princ "\nExcel application released.\n"))
    ));if<-progn
    (if *ExternalDocument* (progn
        (close *ExternalDocument*)
        (setq *ExternalDocument* nil)
        (if (not $bSilent)(princ "\ncsv/txt file released.\n"))
    ));if<-progn
    (princ)
);defun C:ATPort-Release-Apps

(defun C:ATPort-Release-Apps-Force()
    (if (not *ExcelApp*)
        (setq *ExcelApp* (vlax-get-object "Excel.Application"))
    );if
    ; (if (not *ExternalDocument*)
    ;     (setq *ExternalDocument* (open *ExternalDocument*))
    ; );if
    (C:ATPort-Release-Apps)
    (princ)
);defun C:ATPort-Release-Apps-Force

(defun C:ATPort-Reset-Settings()
    (fcnMemoryLocation T)
    (princ "\nMemory settings have been reset to their default values:")
    (princ "\n1) ATIN is set to ")(princ (fcnLoadDefault "C:ATIN"))(terpri)
    (princ "\n2) ATOUT is set to ")(princ (fcnLoadDefault "C:ATOUT"))(terpri)
    (princ "\n3) ATPORT is set to ")(princ (fcnLoadDefault "C:ATPort-With"))(terpri)
    (princ)
);defun C:ATPort-DEFAULT-Reset
(defun C:ATPort-Set (/ sAtriName sNewDefault)
    (initget 327 "ATIN ATOUT")
    (setq sAtriName (getkword "Select which function to change (ATIN ATOUT): "))
    (initget 327 "All Name Select")
    (setq sNewDefault (getkword "Select one of the following options for " 
                                sAtriName " (All Name Select): "))
    (fcnSetDefault (strcat "C:" sAtriName) sNewDefault)
    (princ (strcat "\n" sAtriName " is now set to mimic " sAtriName "-" 
        (fcnLoadDefault (strcat "C:" sAtriName) nil)))(terpri)
    (princ)
);defun C:ATPort-Set
(defun C:ATPort-With (/ sUserSelection)
    (princ (strcat "\nATPort is currently set to port with " (fcnLoadDefault "C:ATPort-With") " files."))
    (princ "\nSelect one of the following options to change which file type attributes are exported into for storage. ")
    (initget 327 "csv Excel txt")
    (setq sUserSelection (getkword "Attribute Export (csv, Excel, txt) : "))
    (fcnSetDefault "C:ATPort-With" (strcase sUserSelection nil))
    (princ (strcat "\nATOUT will now export into " sUserSelection " files."))
    (princ)
);defun C:ATPort-With

;;; User Selection for blocks
(defun fcnBlockSelect (sSelectType / ssBlocks SelectionSet)
    
    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnBlockSelect - Start" 1))

    ;; How does the user want to select the blocks?
    (setq ssBlocks nil)
    (setq SelectionSet (cond
        ((= sSelectType "All");---; Condition 1 - Pulls in all of the blocks found within the document
            (ssget "_X" (list (cons 0 "INSERT")(cons 66 1))))
        ((= sSelectType "Name");--; Condition 2 - User types in a particular block name
            (setq ssBlocks (getstring "\n	***	Block name:\n"))
            (ssget "_X" (list (cons 0 "INSERT")(cons 66 1) (cons 2 ssBlocks)))
        )((= sSelectType "Select"); Condition 3 - User selects the blocks on screen
            (ssget '((0 . "INSERT")))
        )(T nil) ; ---------------; Condition Else
    ));setq->cond

    ;; Error Message - End
    (if *bProgrammerDebug* (fcnBuildError "fcnBlockSelect - End" -1))

    SelectionSet
); defun fcnBlockSelect

;;; NOTICE: These functions run before *error* is assigned to fcnErrorFunction.

;;; User Interactions ================================================== User Interactions ;;;
;;; ====================================================================================== ;;;
;;; Default Settings ==================================================== Default Settings ;;;

;;; Default Values - Pull
(defun fcnLoadDefault (sKey / ;------------; Input Variable
    sPathAndName txtDoc bNextLine sLineInfo ;
    sDefaultValue OrigErrMsg); Local Declarations

    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnLoadDefault - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "An error occured while loading the default value.")

    ;; Initializing
    (setq sKey (fcnToString sKey))
    (setq sPathAndName (fcnMemoryLocation nil))

    (setq txtDoc (open sPathAndName "r"))
    (setq bNextLine T)
    (setq sDefaultValue "")

    ;; Reading default file
    (while (and bNextLine (setq sLineInfo (read-line txtDoc)))
        (if (> (strlen sLineInfo)(strlen (fcnSubstitute "" sKey sLineInfo)))(progn 
            (setq sDefaultValue (fcnSubstitute "" sKey sLineInfo))
            (setq sDefaultValue (fcnSubstitute "" "=" sDefaultValue))
            (setq sDefaultValue (vl-string-trim " " sDefaultValue))
            (setq bNextLine nil)
        ));if<-progn
    );while
    (close txtDoc)

    ;; Validation Warning
    (if bNextLine (progn
        (princ (strcat "Error: The key, \"" sKey "\", was not found in the list."))
        (exit)
    ));if<-progn

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnLoadDefault - End" -1))

    sDefaultValue
);defun fcnLoadDefault

;;; Default Values - Set
(defun fcnSetDefault (sKey sNewValue / ;---; Input Variables
    sPathAndName txtDoc bNoMatch sLineInfo lContents OrigErrMsg
    ); Local Declarations

    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnSetDefault - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured while setting the default value.")

    ;; Initializing Read
    (setq sPathAndName (fcnMemoryLocation nil))
    (setq txtDoc (open sPathAndName "r"))
    (setq bNoMatch T)
    (setq lContents (list))

    ;; Reading default file
    (while (setq sLineInfo (read-line txtDoc))
        (if (> (strlen sLineInfo)(strlen (fcnSubstitute "" sKey sLineInfo)))(progn
            (setq lContents (cons (strcat sKey " = " sNewValue) lContents)); True - Line Matches key word
            (setq bNoMatch nil)
        );progn-True
            (setq lContents (cons sLineInfo lContents));-------------------; False - Copy Line
        );if
    );while
    (close txtDoc)
    
    ;; Initializing Write
    (if bNoMatch (setq lContents (cons (strcat sKey " = " sNewValue) lContents)))
    (setq lContents (reverse lContents))
    (setq txtDoc (open sPathAndName "w"))

    ;; Writing change to the default file
    (foreach sLineInfo lContents
        (write-line sLineInfo txtDoc)
    );while 
    
    ;; Global Variable
    (if (= sKey "C:ATPort-With")(setq *ExternalFileType* sNewValue))

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnSetDefault - End" -1))

    ;; Printing Return
    (princ "\nDefault setting updated\n")
    (princ)
);defun fcnSetDefault

;;; Default Values - Address
(defun fcnMemoryLocation ( $bForceNew / ;-; Input Variables
    sAppDataRoam sPathAndName sPnN txtDoc ; File path, name, and variable
    OrigErrMsg ;--------------------------; Original Error Message
    ); Local Declarations

    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnMemoryLocation - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured while finding the location to store variable information.")

    ;; Initializing
    (setq sAppDataRoam (getenv "AppData")); "C:\\Users\\[User]\\AppData\\Roaming"
    (setq sPathAndName (strcat sAppDataRoam "\\CadAddins\\AT-Import-Export-Defaults.txt"))
    (setq sPnN (findfile sPathAndName))
    (if $bForceNew (setq sPnN ""))

    ;; Creating default file
    (if (/= sPathAndName sPnN)(progn
        ;; Directory Path
        (if (not (findfile (vl-filename-directory sPathAndName)))
            (vl-mkdir (vl-filename-directory sPathAndName))
        );if

        ;; Default settings
        (setq txtDoc (open sPathAndName "w"))
        (write-line "C:ATIN = All" txtDoc)
        (write-line "C:ATOUT = Select" txtDoc)
        (write-line "C:ATPort-With = Excel" txtDoc)
        (write-line "Excel Type = csv" txtDoc)
        (write-line "Excel Delim = ," txtDoc)
        (close txtDoc)
        (setq *ExternalFileType* (fcnLoadDefault "C:ATPort-With"))
    ));if<-progn

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnMemoryLocation - End" -1))

    sPathAndName
);fcnMemoryLocation

;; Temporary File Location - Used for intermediary files
(defun fcnTemporaryFolderLocation (/ sLocalAppData sPath OrigErrMsg)
    
    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnMemoryLocation - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nUnable to find the temporary folder path.")

    (setq sLocalAppData (getenv "LocalAppData")); "C:\\Users\\[User]\\AppData\\Local"
    (setq sPath (strcat sLocalAppData "\\Temp"))

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnMemoryLocation - End" -1))

    sPath
);defun fcnTemporaryFolderLocation

;;; Default Settings ==================================================== Default Settings ;;;
;;; ====================================================================================== ;;;
;;; Main ATOUT ================================================================ Main ATOUT ;;;

(defun fcnMainExport (ssBlock / ;-------------; Input Variable
    lMasterList lTitleList lBodyList ;--------; Inbound List Variables
    lBlock lAttriList lNameValue ;------------; Breaking down inbound list
    sHandle sBlkName sAttriName sAttriValue ;-; List Components
    lExternalList sName sValue ;--------------; Outbound List Variables
    ObjDoc sFileName sFilePath sFileExtension ; File based variables
    sUserFileNameAndPath sFileNameAndPath ;---; File based variables
    iItr1 iItr2 sDelim ;----------------------; Micellaneous
    OrigErrMsg sSaveTo ;----------------------; Error message from parent function
    bCustom); Local Declarations

    ;; Error Handling  
    (setq *OriginErrState* *error*)
    (setq *error* fcnErrorFunction)

    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnMainExport - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured in the main attribute export function.")

    ;; Initializing Variables
    (setq lMasterList (fcnGetAttriMasList ssBlock))
    (setq lTitleList (nth 0 lMasterList))
    (setq lBodyList (nth 1 lMasterList))
    (setq sDelim (cond
        ((= (strcase *ExternalFileType*) "CSV") ",")
        ((= (strcase *ExternalFileType*) "TXT") "\t")
        ((= (strcase *ExternalFileType*) "EXCEL") (fcnLoadDefault "Excel Delim"))
    ));setq<-cond
    
    ;; Title Names
    (if *bProgrammerDebug* (fcnBuildError "fcnMainExport - Titles" 0))
    (setq sValue "")
    (foreach sName lTitleList 
        (setq sValue (strcat sValue sDelim sName))
    );foreach
    (setq sValue (substr sValue (1+ (strlen sDelim))(- (strlen sValue)(strlen sDelim))))
    (setq lExternalList (list sValue))

    ;; Each Block
    (terpri)
    (if *bProgrammerDebug* (fcnBuildError "fcnMainExport - Blocks" 0))
    (foreach lBlock lBodyList
        (setq sHandle (nth 0 lBlock))
        (setq sBlkName (nth 1 lBlock))
        (setq lAttriList (nth 2 lBlock))
        (setq sValue (strcat sHandle sDelim sBlkName))

        ;; Looping through titles
        (setq iItr1 1); Skipping Handle and Blockname
        (while (> (length lTitleList)(setq iItr1 (1+ iItr1)))
            (setq sName (nth iItr1 lTitleList))

            ;; Each Attribute Value
            (setq sAttriValue nil)
            (setq iItr2 -1)
            (while (and (> (length lAttriList)(setq iItr2 (1+ iItr2))) (not sAttriValue))
                (setq lNameValue (nth iItr2 lAttriList))
                (setq sAttriName (nth 0 lNameValue))
                (if (= sAttriName sName)(setq sAttriValue (nth 1 lNameValue)))
            );while

            ;; Saving result
            (if sAttriValue
                (progn
                    (setq sAttriValue (fcnSubstitute "\"\"" "\"" sAttriValue)); Doubles the quote marks
                    (if (/= (strlen sAttriValue) (strlen (fcnSubstitute "" sDelim sAttriValue)))
                        (setq sAttriValue (strcat "\"" sAttriValue "\"")); quote mark's the delimiter
                    );if
                    (setq sValue (strcat sValue sDelim sAttriValue))
                );progn
                (setq sValue (strcat sValue sDelim "<>"))
            );if
        );while
        
        ;; Building master list
        (setq lExternalList (append lExternalList (list sValue)))
    );foreach

    ;; File Extension Type
    (if *bProgrammerDebug* (fcnBuildError "fcnMainExport - File Extension Type" 0))
    (setq bCustom nil)
    (cond
        ;; Condition 1 - Direct
        ((=(strcase *ExternalFileType*) "CSV")
            (setq sFileExtension "csv")
            (setq sSaveTo nil)
            (setq bCustom nil)
        ;; Condition 2 - Direct
        )((=(strcase *ExternalFileType*) "TXT")
            (setq sFileExtension "txt")
            (setq sSaveTo nil)
            (setq bCustom nil)
        ;; Condition 3 - Indirect
        )((=(strcase *ExternalFileType*) "EXCEL")
            (setq sFileExtension (fcnLoadDefault "Excel Type"))
            (setq sSaveTo nil)
            (setq bCustom T)
        ; ;; Condition 4 - Indirect
        ; )((=(strcase *ExternalFileType*) "[Custom File Type]")
        ;     (setq sFileExtension "[File type extension]")
        ;     (setq sSaveTo "[csv/txt file type]")
        ;     (setq bCustom T)
        ;; Else - Error
        )(T 
            (setq *sErrorMessage* (strcat "Error: The extension type, " (fcnToString *ExternalFileType*) ", in the variable *ExternalFileType* was not found in the list."))
            (if *bProgrammerDebug* (setq *sDebugMessage* (strcat "\nProgrammer: Add/revise relevant information to \"File Extension Type\" under the function, \"fcnMainExport\" in  section, \"Main ATOUT.\"")))
            (exit)
        ); Else
    );cond

    ;; Pulling the file name and changing extension name
    (if *bProgrammerDebug* (fcnBuildError "fcnMainExport - File Path" 0))
    (setq ObjDoc (vla-get-activedocument (vlax-get-acad-object)))
    (setq sFilePath (vl-filename-directory (vla-get-fullname ObjDoc)))
    (setq sFileName(vl-filename-base (vla-get-name ObjDoc)))
    (setq sFileName (strcat sFilePath "\\" sFileName "." sFileExtension))

    ;; User setting the name and path
    (setq sUserFileNameAndPath (getfiled "Output File" sFileName sFileExtension 1))

    ;;File selected
    (if *bProgrammerDebug* (fcnBuildError "fcnMainExport - Selected File" 0))
    (if sUserFileNameAndPath (progn
        ;; CVS/TXT/Other file type save
        (cond 
            ;; Direct
            ((=(strcase *ExternalFileType*) "TXT") (setq sFileNameAndPath sUserFileNameAndPath))
            ((=(strcase *ExternalFileType*) "CSV") (setq sFileNameAndPath sUserFileNameAndPath))
            
            ;; Indirect - Opens csv/txt file with a different application
            ((Not sSaveTo)(setq sFileNameAndPath sUserFileNameAndPath))
            
            ;; Indirect - Creates a document with the specified, custom file type
            (T ; Else
                (setq sFilePath (fcnTemporaryFolderLocation));-----------------------; Sets the intermediary file to the temporary folder
                ;(setq sFilePath (vl-filename-directory sUserFileNameAndPath));-------; Pulls the user defined file path
                (setq sFileName (vl-filename-base sUserFileNameAndPath));------------; Pulls the user defined file name
                (setq sFileNameAndPath (strcat sFilePath "\\" sFileName "." sSaveTo)); Creates a csv/txt file path
            );Else
        );cond

        ;; Writing to csv/txt file
        (fcnWriteToDocument lExternalList sFileNameAndPath)
        
        ;; Indirect Save
        (if bCustom (progn 
            ;; Saving new Excel file
            (if (=(strcase *ExternalFileType*) "EXCEL")(fcnOpenInExcel sFileNameAndPath))
            ;(if (=(strcase *ExternalFileType*) "...")(fcnDocPopulate[AppName] sUserFileNameAndPath sFileNameAndPath))
            ;(if (=(strcase *ExternalFileType*) "...")(fcnDocPopulate[AppName] sUserFileNameAndPath sFileNameAndPath))
            ;(if (=(strcase *ExternalFileType*) "...")(fcnDocPopulate[AppName] sUserFileNameAndPath sFileNameAndPath))
        ));if<-progn
    )(progn ; No file path, no file name
        (princ "\nFile not create.")(terpri)
    ));if<-progn

    ;; Error Message - End
    (princ "\nAttributes successfully exported.\n")
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnMainExport - End" -1))
    (princ)
);defun fcnMainExport

;;; Main ATOUT ================================================================ Main ATOUT ;;;
;;; ====================================================================================== ;;;
;;; Main ATIN ================================================================== Main ATIN ;;;

(defun fcnMainImport (ssBlocks / ; Input Variable
        lRawInput lTitles lFinishedInput ;--------------; List variables to convert
        iItr1 iItr2 ;-----------------------------------; Iterations
        lNameValue lNewAttris lNewBlock lMasterList ;---; List variables
        OrigErrMsg ;------------------------------------; Error message from parent function
        lAttriValues sBlkHandle sBlkName ;--------------; Block variables
        sAttriValue sAttriTag ;-------------------------; Attribute variables
        DSDoc sErrDoc ;---------------------------------; Document variables
        lErrors lStatus ;-------------------------------; Error Lists
        sFound1 sFound2 sFound3 sFound4 sCombineStr ;---; Found/Missing Strings
        iErrBSource iErrBDest iErrAttSource iErrAttDest ; Error Counters
        sRawDateAndTime sDateAndTime ;------------------; Date and time
    ); Local Variables

    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnMainImport - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured in the main attribute import function.")

    ;; Error Handling
    (setq *OriginErrState* *error*)
    (setq *error* fcnErrorFunction)
    
    ;; Getting file from user
    (setq *sFileName* (getfiled "Open CSV or Text delimited data file" 
        (if *sFileName*  *sFileName* "P:\\ENG") "" 0));----------------; file already in memory
    
    ;; Initializing variables
    (setq lRawInput (fcnMakeItemList *sFileName* nil))
    (setq lTitles (nth 0 lRawInput))
    (setq lMasterList (list))
    
    ;; Each Block
    (setq iItr1 0)
    (while (> (length lRawInput) (setq iItr1 (1+ iItr1)))

        ;; Block Details
        (setq lAttriValues (nth iItr1 lRawInput))
        (setq sBlkHandle (nth 0 lAttriValues))
        (setq sBlkName (nth 1 lAttriValues))

        ;; Each Attribute
        (setq iItr2 1)
        (setq lNewAttris (list))
        (while (> (length lAttriValues) (setq iItr2 (1+ iItr2))) 

            ;; Valid Entry
            (setq sAttriValue (nth iItr2 lAttriValues))
            (if (/= sAttriValue "<>")(progn
                (setq sAttriTag (nth iItr2 lTitles))
                (setq lNameValue (list sAttriTag sAttriValue))
                (setq lNewAttris (append lNewAttris (list lNameValue)))
            ));if<-progn
        );while
        (setq lNewBlock (list sBlkHandle sBlkName lNewAttris))
        (setq lMasterList (append lMasterList (list lNewBlock)))
    );while
    
    ;; Updating Block Attributes
    (setq DSDoc (vla-get-activedocument (vlax-get-acad-object)))
    (vla-endundomark DSDoc)
    (vla-startundomark DSDoc)
    (setq lErrors (fcnModBlockAtt lMasterList ssBlocks))
    (vla-endundomark DSDoc)


    ;; Error Reports
    (if *bProgrammerDebug* (fcnBuildError "fcnMainImport - Building Error Report" 0))
    
    ;; Initializing Error Variables
    (setq iErrBSource 0)
    (setq iErrBDest 0)
    (setq iErrAttSource 0)
    (setq iErrAttDest 0)
    (setq iItr1 0)

    ;; Creating Log Document
    (setq sRawDateAndTime (rtos (getvar "CDATE") 2 6))
    (setq sDateAndTime (strcat (substr sRawDateAndTime 1 4)))
    (setq sDateAndTime (strcat sDateAndTime "-" (substr sRawDateAndTime  5 2)))
    (setq sDateAndTime (strcat sDateAndTime "-" (substr sRawDateAndTime  7 2)))
    (setq sDateAndTime (strcat sDateAndTime "_" (substr sRawDateAndTime 10 2))) 
    (setq sDateAndTime (strcat sDateAndTime "-" (substr sRawDateAndTime 12 2)))
    (setq sDateAndTime (strcat sDateAndTime "-" (substr sRawDateAndTime 14 2)))
    (setq sErrDoc (strcat (fcnTemporaryFolderLocation) "\\CADAttInLog - " sDateAndTime ))
    (setq sErrDoc (strcat sErrDoc " - " (vl-filename-base (vla-get-name DSDoc)) ".csv"))
    (setq *ExternalDocument* (open sErrDoc "w"))

    ;; Reporting Errors
    (foreach lStatus lErrors

        ;; Blocks
        (setq sBlkHandle (nth 0 lStatus))
        (setq sBlkName (nth 1 lStatus))
        (setq sFound1 (nth 2 lStatus))
        (setq sFound2 (nth 3 lStatus))

        ;; Attributes
        (setq sAttriTag (nth 4 lStatus))
        (setq sFound3 (nth 5 lStatus))
        (setq sFound4 (nth 6 lStatus))

        ;; Counters
        (if (= "Missing" sFound1)(setq iErrBSource (1+ iErrBSource)))
        (if (= "Missing" sFound2)(setq iErrBDest (1+ iErrBDest)))
        (if (= "Missing" sFound3)(setq iErrAttSource (1+ iErrAttSource)))
        (if (= "Missing" sFound4)(setq iErrAttDest (1+ iErrAttDest)))
        (if (= "Found" sFound1 sFound2 sFound3 sFound4)(setq iItr1 (1+ iItr1)))

        ;; Writing error report to document
        (setq sCombineStr (strcat "\"" sBlkHandle "\", \"" sBlkName))
        (setq sCombineStr (strcat sCombineStr "\", \"" sFound1 "\", \"" sFound2))
        (setq sCombineStr (strcat sCombineStr "\", \"" sAttriTag))
        (setq sCombineStr (strcat sCombineStr "\", \"" sFound3 "\", \"" sFound4 "\""))
        (write-line sCombineStr *ExternalDocument*)
    );foreach
    
    ;; Closing document
    (close *ExternalDocument*)
    (setq *ExternalDocument* nil)
    (if *bProgrammerDebug* (fcnBuildError "fcnMainImport - Error Report Built" 0))
    
    ;; Displaying results to the user
    (terpri)
    (if (> iErrAttDest 0)  (princ "Warning: Some of the matching blocks in the supplied list don't have matching attributes in the active document.\n"))
    (if (> iErrAttSource 0)(princ "Warning: Some of the matching blocks in the active document don't have matching attributes from the supplied list.\n"))
    (if (> iErrBDest 0)  (princ (strcat "Warning: " (itoa iErrBDest)   " block(s) from the specified file were not found in the active document.\n")))
    (if (> iErrBSource 0)(princ (strcat "Warning: " (itoa iErrBSource) " block(s) in the active document were not found in the supplied file.\n")))
    (princ (strcat "Complete: " (itoa iitr1) " of " (itoa (1- (length lErrors))) " block(s) were successfully updated.\n"))
    (if (not (>= 0 iErrAttDest iErrAttSource iErrBDest iErrBSource))
        (princ (strcat "The detailed error log can be found here: \n" sErrDoc "\n"))
    );if

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnMainImport - End" -1))

    ;; Error Handling
    (setq *error* *OriginErrState*)

    (princ)
);defun fcnMainImport

;;; Main ATIN ================================================================== Main ATIN ;;;
;;; ====================================================================================== ;;;
;;; Get/Pull Block Attributes ================================== Get/Pull Block Attributes ;;;

(defun fcnGetAttriMasList (SelectionSet / ;-; User Input
    iItr1 iMax ;---------------------------; Number variables
    ssEName sHandle sBlockName ;-----------; Selection Set and String variables
    objBlock lBlockAttri lBlocksWithAttris ; Creating Attribute Lists per Block
    lAttNames AttList OrigErrMsg sProgress
    AttName AttDup bAttNew lMasterList
    ); Miscellaneous
    
    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnGetAttriMasList - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured while checking each block.")
    (if *bProgrammerDebug* (setq *sErrorMessage* (strcat *sErrorMessage* "\n")))

    ;; Initializing Vriables
    (setq iItr1 -1)
    (setq iMax (sslength SelectionSet))
    (setq lBlocksWithAttris nil)
    (setq sBlockName nil)
    (setq lAttNames (list "BLOCKNAME" "HANDLE")); order is reversed after while loop

    ;; Creating Attribute Lists per Block
    (terpri); Progress notification
    (while (> iMax (setq iItr1 (1+ iItr1)))

        ;; Creating Attribute Lists per Block
        (setq ssEName (ssname SelectionSet iItr1));-------; Single Entity
        (setq sHandle (cdr (assoc 5 (entget ssEName))));--; Block's Handle
        (setq sBlockName (cdr (assoc 2 (entget ssEName)))); Block's Name
        (setq objBlock (vlax-ename->vla-object ssEName));-; Converts entity into object
        (setq lBlockAttri (fcnGetAttributes objBlock));---; List of attributes from object
        (setq lBlocksWithAttris (append lBlocksWithAttris  
            (list (list sHandle sBlockName lBlockAttri)))); Combining Informtion into one list 
        
        ;; Creating List of attribute names
        (foreach AttList lBlockAttri
            (setq AttName (nth 0 AttList))
            (setq bAttNew T)
            (setq AttDup nil)
            (foreach AttDup lAttNames ; Checks for duplicates
                (if (= AttDup AttName) (setq bAttNew nil)));Dupliate found
            (if bAttNew (setq lAttNames (cons AttName lAttNames))); Adds new, unique name
        );foreach

        ;; Progress notification
        (setq sProgress (strcat "\rWorking on block " (rtos (1+ iItr1) 2 0) " of " (rtos iMax 2 0) "."))
        (princ sProgress)
        (if *bProgrammerDebug* (setq *sErrorMessage* (strcat *sErrorMessage* sProgress)))

        ; ;; Debug Print
        ; (princ (strcat "\nBlock " (rtos (1+ iItr1) 2 0) " out of " (rtos iMax 2 0) "\n"))
        ; (princ "objBlock : ")(princ objBlock)(terpri)
        ; (princ "Handle : ")(princ (cdr (assoc 5 (entget ssEName))))(terpri)
        ; (princ "Block Name : ")(princ (cdr (assoc 2 (entget ssEName))))(terpri)
        ; (princ "lBlockAttri : ")(princ lBlockAttri)(terpri)
        ; (princ "=============================================\n")
    );while
    (princ "..Completed\n"); Progress notification
    (if *bProgrammerDebug* (setq *sErrorMessage* (strcat *sErrorMessage* "..Completed")))
    (setq lAttNames (reverse lAttNames))

    ; ;; Debug Print Complete Attribute List
    ; (princ "\nAttributes Collected\n")
    ; (princ "\n \nlAttNames : \n")(princ lAttNames)
    ; (princ "\n \nlBlocksWithAttris : \n")(princ lBlocksWithAttris)(terpri)
    ; (princ "\n \nAttributes Collected\n")
    ; (princ "=============================================\n")
    (setq lMasterList (list lAttNames lBlocksWithAttris))
    ;;-------------------------------------------------------------------------------------;;

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnGetAttriMasList - End" -1))
    
    lMasterList
);defun fcnGetAttriMasList

;; Local function
(defun fcnGetAttributes (objBlock / 
    ssAttributes ssAtt attTag attValue attList OrigErrMsg)
    
    ;; Error Message - Start
    (setq OrigErrMsg *sErrorMessage*)
    (if *bProgrammerDebug* (fcnBuildError "fcnGetAttributes - Start" 1))
    (setq *sErrorMessage* "\nAn error occured while collecting attributes from a block.")

    ;; Validation
    (if (not objBlock)(progn (princ "\nNo object provided.\n") (exit))); No object

    ;; Initializing
    (setq attList nil)
    (setq ssAttributes (vlax-invoke objBlock 'GetAttributes))
    (setq ssAttributes (vlax-safearray->list ssAttributes))
    (setq ssAttributes (mapcar 'vlax-variant-value ssAttributes))
    
    ;; Creating list
    ; (princ "\n-------------------------------------\n")
    (if *bProgrammerDebug* (setq *sErrorMessage* (strcat 
        *sErrorMessage* "\n-------------------------------------")))
    (foreach ssAtt ssAttributes
        (setq attTag (vla-get-tagstring ssAtt))
        (setq attValue (vla-get-textstring ssAtt))
        (setq attList (append attList (list (list attTag attValue))))

        ; ;; Debug Print
        (if *bProgrammerDebug* (setq *sErrorMessage* (strcat 
            *sErrorMessage* "\nattTag : " (fcnToString attTag)
                            "\nattValue : " (fcnToString attValue)
                            "\n-------------------------------------\n"
        )));if<-setq<-strcat
        ; (princ (strcat "ssAtt : "))    (princ ssAtt)    (terpri)
        ; (princ (strcat "attTag : "))   (princ attTag)   (terpri)
        ; (princ (strcat "attValue : ")) (princ attValue) (terpri)
        ; (princ "-------------------------------------\n")
    );while
    
    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnGetAttributes - End" -1))
    
    attList
);defun fcnGetAttributes

;;; Get/Pull Block Attributes ================================== Get/Pull Block Attributes ;;;
;;; ====================================================================================== ;;;
;;; Set/Push Block Attributes ================================== Set/Push Block Attributes ;;;

(defun fcnModBlockAtt (lMasterList ssBlocks / ; Input Variables
        ssEName objBlock ;-----------------------; Block variables for within the drawing
        sHandle sBlockName slHandle slBlockName ;; String Variables
        iItr1 iItr2 iItr3 iMaxss iMaxml iMaxAtt
        OrigErrMsg sReturnError lReturnError
        lBlockList bFindNext lAttList lAttNameValue
        lErrorBlock lAttributeState
        ); Local Variable Declarations
    ;; 
    ;; sl[Variable] = String variable from Master List (String List [Variable])
    
    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnModBlockAtt - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured while updating the specified blocks.")

    ; ;; Selection set of blocks
    ; (setq ssBlocks (ssget "x" (list (cons 0 "INSERT"))))

    ;; Initializing Vriables
    (setq iItr1 -1)
    (setq iMaxss (sslength ssBlocks))
    (setq iMaxml (length lMasterList))
    (setq sBlockName nil)
    (setq lReturnError (list (list "Handle" "Blockname" "Block Source" "Block Destination" 
                            "Attribute Name" "Attribute Source" "Attribute Destination")))

    ;; Each selected block in the active document
    (while (> iMaxss (setq iItr1 (1+ iItr1)))
        
        ;; Finding the Handle and BlockName for the block
        (setq ssEName (ssname ssBlocks iItr1));-----------; Single Entity
        (setq sHandle (cdr (assoc 5 (entget ssEName))));--; Block's Handle
        (setq sBlockName (cdr (assoc 2 (entget ssEName)))); Block's Name

        ;; Each block in the Master List
        (setq iItr2 -1)
        (setq bFindNext T)
        (while (and (> iMaxml (setq iItr2 (1+ iItr2))) bFindNext)
            
            (setq lBlockList (nth iItr2 lMasterList))
            (setq slHandle (car lBlockList))
            (setq slBlockName (cadr lBlockList))

            ;; Block Match
            (if (and (= sHandle slHandle)(= sBlockName slBlockName))(progn
                (setq objBlock (vlax-ename->vla-object ssEName)); Converts entity into object
                (setq lAttList (nth 2 lBlockList));-------------; Pulling attributes from block list from master list
                (setq lErrorBlock (list sHandle sBlockName "Found" "Found"))

                ;; Each Attribute
                (setq iItr3 -1)
                (setq iMaxAtt (length lAttList))
                (while (> iMaxAtt (setq iItr3 (1+ iItr3)))
                    (setq lAttNameValue (nth iItr3 lAttList));----------------------------------; Name and value
                    (setq lAttributeState (fcnSetAttributes objBlock lAttNameValue));-----------; Set/Pushing Block Attributes
                    (setq lReturnError (cons (append lErrorBlock lAttributeState) lReturnError)); Error messages
                );while
                (setq bFindNext nil)
            ));if<-progn
        );while

        ;; Missing Source Blocks
        (if bFindNext (progn
            (setq lErrorBlock (list sHandle sBlockName "Missing" "Found" "#N/A" "#N/A" "#N/A"))
            (setq lReturnError (cons lErrorBlock lReturnError))
        ));if<-progn
    );while

    ;; Missing Destination Blocks
    ;; Each Block in the Master List
    (setq iItr1 -1)
    (while (> iMaxml (setq iItr1 (1+ iItr1)))
        (setq lBlockList (nth iItr1 lMasterList))
        (setq slHandle (car lBlockList))
        (setq slBlockName (cadr lBlockList))

        ;; Each selected block in the active document
        (setq iItr2 -1)
        (setq bFindNext T)
        (while (and (> iMaxss (setq iItr2 (1+ iItr2))) bFindNext)
            (setq ssEName (ssname ssBlocks iItr2));-----------; Single Entity
            (setq sHandle (cdr (assoc 5 (entget ssEName))));--; Block's Handle
            (setq sBlockName (cdr (assoc 2 (entget ssEName)))); Block's Name
            
            ;; Attribute Check
            (if (and (= sHandle slHandle)(= sBlockName slBlockName))(progn
                (setq lErrorBlock (list sHandle sBlockName "Found" "Found"));----; Sets the first part of the error sequence
                (setq lErrorBlock (fcnAttriCheck ssEName lBlockList lErrorBlock)); Compares attributes and returns completed list
                (setq bFindNext nil);--------------------------------------------; Block match found
            ));if<-progn
        );while
        (if bFindNext (progn
            (setq lErrorBlock (list slHandle slBlockName "Found" "Missing" "#N/A" "#N/A" "#N/A"))
            (setq lReturnError (cons lErrorBlock lReturnError))
        ));if<-progn
    );while
    
    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnModBlockAtt - End" -1))

    (reverse lReturnError)
);defun fcnModBlockAtt

;; Local function
(defun fcnSetAttributes (objBlock lNameValue / ; Input Variables
    sName sValue ;-----------------------------; values from the list input variable
    lAttributes ObjAtt attTag attList ;--------; Block Attributes
    iItr1 iMax OrigErrMsg ;--------------------; Miscellaneous
    bFindNext attValue lReturnError
    ); Variable Declarations
    
    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnSetAttributes - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured while updating the attributes with the specified blocks.")

    ;; Initializing
    (setq sName (nth 0 lNameValue));; Matching Name
    (setq sValue (nth 1 lNameValue)); New value
    (setq lAttributes (vlax-invoke objBlock 'GetAttributes)); Exsisting attributes
    (setq lAttributes (vlax-safearray->list lAttributes))
    (setq lAttributes (mapcar 'vlax-variant-value lAttributes))

    (setq iItr1 -1)
    (setq iMax (length lAttributes))
    (setq attList nil)
    (setq bFindNext T)

    ;; Checking Each Attribute Name
    (while (and (> iMax (setq iItr1 (1+ iItr1))) bFindNext)
        
        ;; Incrementing
        (setq ObjAtt (nth iItr1 lAttributes));-----; Exsisting Attribute
        (setq attTag (vla-get-tagstring ObjAtt));--; Exsisting Name
        (setq attValue (vla-get-textstring ObjAtt)); Exsisting Value
        
        ;; Set/Pushing Tags
        (if (= attTag sName)(progn
            (vla-put-textstring ObjAtt sValue)
            (setq bFindNext nil)
        ));if
    );while

    ;; Error Message
    (if bFindNext
        (setq lReturnError (list sName "Found" "Missing"))
        (setq lReturnError (list sName "Found" "Found"))
    );if

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnSetAttributes - End" -1))

    ;; Returned value
    lReturnError
);defun fcnSetAttributes

;; Local function
(defun fcnAttriCheck (ssEName lBlockList lStartingError / 
        objBlock lAttributes ObjAtt attTag ; Active document based variables
        lAttList lNameValue sName ;--------; Master list based variables
        iItr1 iItr2 iMaxObj iMaxList ;-----; Itteration based variables
        bFindNext ;------------------------; Match found
        lReturnError lErrorBlock ;---------; Missing matches - Return list
        OrigErrMsg ;-----------------------; Error message
    ); Local Declarations

    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnAttriCheck - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured while capturing attributes without matches.")

    ;; Initializing Blocks
    (setq objBlock (vlax-ename->vla-object ssEName));-------; Converts entity into object
    (setq lAttributes (vlax-invoke objBlock 'GetAttributes)); Exsisting attributes
    (setq lAttributes (vlax-safearray->list lAttributes))
    (setq lAttributes (mapcar 'vlax-variant-value lAttributes))
    (setq iMaxObj (length lAttributes))
    (setq iItr1 -1)

    ;; Initializing List
    (setq lAttList (nth 2 lBlockList)); Pulling attributes from block list
    (setq iMaxList (length lAttList))

    ;; Checking Each Attribute Name
    (while (and (> iMaxObj (setq iItr1 (1+ iItr1))))
        (setq ObjAtt (nth iItr1 lAttributes));-----; Exsisting Attribute
        (setq attTag (vla-get-tagstring ObjAtt));--; Exsisting Name
        (setq bFindNext T)
        
        ;; Checking for a match
        (setq iItr2 -1)
        (while (and (> iMaxList (setq iItr2 (1+ iItr2))));
            (setq lNameValue (nth iItr2 lAttList));------;
            (setq sName (nth 0 lNameValue));-------------; Matching Name
            (if (= attTag sName)(setq bFindNext nil))
        );while
        
        ;; No match
        (if bFindNext (progn
            (setq lErrorBlock (append lStartingError (list attTag "Missing" "Found")))
            (setq lReturnError (append lReturnError (list lErrorBlock)))
        ));if
    );while

        ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnAttriCheck - End" -1))

    lReturnError
);defun fcnAttriCheck

;;; Set/Push Block Attributes ================================== Set/Push Block Attributes ;;;
;;; ====================================================================================== ;;;
;;; Set/Push External Document ================================ Set/Push External Document ;;;

;; cvs/txt Document
(defun fcnWriteToDocument (lInputList sFileName / ; Input Variables
        sContents sOrigErrMsg 
    ); Local Declarations
    
    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnWriteToDocument - Start" 1))
    ;; Display the same message to the user and the programmer
    (setq sOrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* (strcat "\nError: Unable to write to the document \"" sFileName ".\""))

    ;; Initializing
    (setq *ExternalDocument* (open sFileName "w"))
    (if (and *ExternalDocument* lInputList)(progn
        ;; Writing
        (foreach sContents lInputList
            (write-line sContents *ExternalDocument*)
        );foreach

        ;; Finalizing
        (close *ExternalDocument*)
        (setq *ExternalDocument* nil)
    ));if<-progn

    ;; Error Message - End
    (setq *sErrorMessage* sOrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnWriteToDocument - End" -1))
);defun fcnWriteToDocument

;; Open in Excel
(defun fcnOpenInExcel ($CSVFile / ;-; Input Variables
    bFileExsists objWorkbooks bFileClosed Workbook ;; Excel workbook variables
    Sheets ishMax Worksheet ;-----------------------; Excel worksheet variables
    lMasterList iColCount sDelim ;------------------; CVS file variables
    sQueryFormula ;---------------------------------; 
    iIterMax iIter1 OrigErrMsg ;--------------------; 
    bNewSheet objQueries iItr1 bSearchForDup sTitle ;
    objCellA1 objListObjects objTableQuery ;--------;
    objQuery
    ); Local Declarations

    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnOpenInExcel - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error has occured in function to open the file with Excel.")
    (if (= (type $CSVFile) 'STR)
        (if (not (findfile $CSVFile))(progn
            (alert (strcat "File not found:\n" $CSVFile))
            (exit)
        ));if
        (progn
            (alert (strcat "Source file path not specified."))
            (exit)
        );progn
    );if

    ;; Assigning Excel Application
    (if (not *ExcelApp*)(setq *ExcelApp* (vlax-get-or-create-object "Excel.Application")))
    
    ;; Opening the Excel workbook
    (setq Workbook (vlax-invoke-method (vlax-get-property *ExcelApp* 'WorkBooks) 'Open $CSVFile))

    (vlax-release-object *ExcelApp*)(gc);--------------------; Releasing from memory
    (setq *ExcelApp* nil);-----------------------------------; Setting variable to nothing

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnOpenInExcel - End" -1))
    (princ)
);defun fcnOpenInExcel

;;; Set/Push External Document ================================ Set/Push External Document ;;;
;;; ====================================================================================== ;;;
;;; Get/Pull External Document ================================ Get/Pull External Document ;;;

;; cvs/txt Document
(defun fcnMakeItemList ( ; Input Variables
        $sThisFile $bByColumn ; Returns output in a "By Column" format instead of a "By Row" format
        ;-; Input Variables 
        /
        ;-; Local Declarations 
        iItr1 ;--------; Iteration variable
        sNextLine ;----; Line of text in document
        lMappedList ;--; List of values (by column)
        lReadList ;----; List of text in document
        CharCnt ;------; Number of characters in a string
        lDataList ;----; List of values (by row)
        lTitleList ;---; List of Title variables (First row)
        lReturnList ;--; Returns the list
        sWCMatchDelim ;; Delimiter
        bCheckEnd ;----; Extra Delimeter at the end of the string
        bTextDelim ;---; Ignore found delimeter - Part of string segment
        bToBeTextDelim ; Ignore the next delimeter found
        bLastTextDelim ; Last delimeter was ignored - Run special capture condition on string segment
        OrigErrMsg ;---; Contains any errors generated from any previous functions. 
    ); Local Declarations

    ;; Error Message - Start
    (setq OrigErrMsg *sErrorMessage*)
    (if *bProgrammerDebug* (progn 
        (fcnBuildError "fcnMakeItemList - Start" 1)
        (setq *sErrorMessage* "")
    )(progn
        (setq *sErrorMessage* "\nError: An issue occured while pulling the attributes from the selected file.")
    ));if<-progn

    (progn ;Routine's Info
        ;-------------------------------------------------------------------------------
        ; Origin: https://forums.autodesk.com/t5/visual-lisp-autolisp-and-general/autolisp-for-excel-to-separate-lists/td-p/6978304
        ; Program Name: C:fcnMakeItemList
        ; Created By:   Cooper Francis
        ; Modified By:  Garrett Beck (06/02/2022)
        ; Product Ver.: 13.4.1385.0 Autodesk Civil 3D 2022.1.3 Update
        ; Built on:        S.162.0.0; AutoCAD 2022.1.2
        ;                        25.0 45.5 Autodesk AutoCAD Map 3D 2022.0.1
        ;                        8.3.53.0 AutoCAD Architecture 2022 
        ;-------------------------------------------------------------------------------
    )

    ;; Initializing data list
    (setq lDataList nil)
    (setq lReturnList nil)

    ;; Collecting from file
    (if $sThisFile (progn ; Does file name exsists
        (setq *ExternalDocument* (open $sThisFile "r"))
        (if *ExternalDocument* (progn ; Did file open
            
            (if *bProgrammerDebug* (setq *sErrorMessage* (strcat "\nfcnMakeItemList - Document \"" $sThisFile "\" opened.")))

            ;; Initializing Variables
            (setq iItr1 0)
            (setq CharCnt 0)
            (cond ; File Type
                ((wcmatch $sThisFile "*`.csv")(setq sWCMatchDelim "`,")); File type: csv
                ((wcmatch $sThisFile "*`.txt")(setq sWCMatchDelim "\t")); File type: txt
                (T 
                    (princ "\nInvalid file extension selected.")
                    (exit)
                );Else
            );cond
            
            ;; Updates sNextLine with the next line of text in the document
            (while (setq sNextLine (read-line *ExternalDocument*))
                (setq sNextLine (fcnsubstitute "\"" "\"\"" sNextLine))

                ;; Collecting Title values
                (setq iItr1 (1+ iItr1))

                ;; Error Message
                (if *bProgrammerDebug* (setq *sErrorMessage* (strcat *sErrorMessage* "\nfcnMakeItemList - " (rtos iItr1 2 0) ") sNextLine : " sNextLine)))

                ;; Converts separator into list
                ;; List of items in line, split with "sWCMatchDelim"
                (setq lReadList (list))
                (setq bTextDelim nil)
                (setq bLastTextDelim nil)
                (setq CharCnt 0)
                (while (wcmatch (substr sNextLine CharCnt) (strcat "*" sWCMatchDelim "*"))
                    (if (wcmatch (substr sNextLine CharCnt 2) (strcat sWCMatchDelim "\""))(setq bToBeTextDelim T))
                    (if (wcmatch (substr sNextLine CharCnt 2) (strcat "\"" sWCMatchDelim))(setq bTextDelim nil))
                    (if (and (wcmatch (substr sNextLine CharCnt) (strcat sWCMatchDelim "*")) (not bTextDelim))(progn
                        (if bLastTextDelim
                            (setq lReadList (append lReadList (list (substr sNextLine 2 (- CharCnt 3))))) 
                            (setq lReadList (append lReadList (list (substr sNextLine 1 (- CharCnt 1)))))
                        );if
                        (setq sNextLine (substr sNextLine (1+ CharCnt))) 
                        (setq CharCnt 0)
                        (setq bTextDelim bToBeTextDelim)
                        (setq bLastTextDelim bToBeTextDelim)
                        (setq bToBeTextDelim nil)
                    ));if<-progn
                    (setq CharCnt (1+ CharCnt))
                );while
                (if bLastTextDelim
                    (setq lReadList (append lReadList (list (substr sNextLine 2 (- (strlen sNextLine) 2)))))
                    (setq lReadList (append lReadList (list (substr sNextLine 1))))
                );if

                ;; Store data in either Title or body
                (cond
                    ;; Condition 1 - Title Section
                    ((= iItr1 1)
                        (if (= 0 (strlen (car (reverse lReadList))))(progn ; Checks for extra delimiter
                            (setq bCheckEnd T); Check the lines underneath the title for extra delimiter
                            (setq lReadList (reverse (cdr (reverse lReadList)))); Removes extra delimiter
                        ));if<-progn
                        (setq lTitleList lReadList)
                    ); Condition 1
                    ;; Condition 2 - Body Section
                    ((> iItr1 1)
                        (if (and bCheckEnd (= 0 (strlen (car (reverse lReadList))))); Removes extra delimiter, if title has extra delimiter
                            (setq lReadList (reverse (cdr (reverse lReadList))))); Removes extra delimiter
                        (setq lDataList (append lDataList (list lReadList)))
                    ); Condition 2
                );cond
            );while
            (if *bProgrammerDebug* (setq *sErrorMessage* (strcat *sErrorMessage* "\nfcnMakeItemList - Closing document.")))

            ;; Closes the opened document
            (close *ExternalDocument*)(gc)
            (setq *ExternalDocument* nil)
            
            ;; Converts data from "by row" to "by column"
            (if *bProgrammerDebug* (setq *sErrorMessage* (strcat *sErrorMessage* "\nfcnMakeItemList - Organizing return list.")))
            (if (and lTitleList lDataList $bByColumn)
                (progn ; True - Return by Column
                    (setq iItr1 0)
                    (setq lReturnList (list))
                    (while (< iItr1 (length lTitleList))
                        (setq lMappedList (mapcar '(lambda (x) (nth iItr1 x)) lDataList))
                        (while (or (not (last lMappedList))(eq (last lMappedList) ""))
                            (setq lMappedList (reverse (cdr (reverse lMappedList))))
                        );while
                        (setq lReturnList (append lReturnList (list (list (nth iItr1 lTitleList) lMappedList))))
                        (setq iItr1 (1+ iItr1))
                    );while
                )(progn ; False - Return by Row
                    (setq lReturnList (list lTitleList))
                    (setq lReturnList (append lReturnList lDataList))
            ));if<-progn
        );progn
            ;; Else - *ExternalDocument* was nil
            (if $sThisFile
                (princ (strcat "\nUnable to open \"" $sThisFile "\" for reading! "))
                (princ (strcat "\nNo file was given to open! "))
            );if 
        );if
    );progn
        ;; Else - file path not selected
        (princ "\nNo CSV file was selected! ")
    );if

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnMakeItemList - End" -1))
    lReturnList
); defun fcnMakeItemList

;;; Get/Pull External Document ================================ Get/Pull External Document ;;;
;;; ====================================================================================== ;;;
;;; Other Functions ====================================================== Other Functions ;;;

;-------------------------------------------------------------------------------
; fcnSubstitute - Substitutes every occurance of a pattern within a string with 
;                   a new substring. If an invalid value is entered, then 
;                   sString is returned.
; Function By: Garrett Beck from Hot Springs, AR, United States
; Arguments: 3
;   sNewPattern = The new pattern for the string
;   sPattern    = The exsisting pattern within the string
;   sString     = The original string in which to Set/Push
; Syntax example: (Substitutes "uu" "u" "Substitutes") = "Suubstituutes"
;Notes:         vl-string-subst substitutes only the first occurrence it finds in the string
;vl-string-subst Source: http://docs.autodesk.com/ACD/2014/ENU/files/GUID-D8EE91DC-D4DB-43E0-9AFE-5FA166C0896D.htm
;-------------------------------------------------------------------------------
;; Substitute every occurrence
(defun fcnSubstitute ( sNewPattern sPattern sString / 
        lString sBuildString sNewString sReturn OrigErrMsg iItr1)

    ;; Validation and default return
    (setq sReturn sString); Unchanged return value
    (if (/= (type sString) 'STR)(setq sString (fcnToString sString)))
    (if (/= (type sPattern) 'STR)(setq sPattern (fcnToString sPattern)))
    (if (/= (type sNewPattern) 'STR)(setq sNewPattern (fcnToString sNewPattern)))

    ;; Error Message - Start
    (setq OrigErrMsg *sErrorMessage*)
    (if *bProgrammerDebug* (progn
        (fcnBuildError "fcnSubstitute - Start" 1)
        (setq *sErrorMessage* (strcat 
                "\nsNewPattern : " sNewPattern 
                "\nsPattern : " sPattern
                "\nsString : " sString
        ));setq<-strcat
    )(progn
        (setq *sErrorMessage* "Error: An error has occured while substituting strings values.")
    ));if

    ;; Validating case
    (if (and (>= (strlen sString)(strlen sPattern))(> (strlen sPattern) 0))(progn
        
        ; (princ "\n--------------------------------------------\n")
        ; (princ "\nsString : ")(princ sString)(terpri)
        ; (princ "\nsPattern : ")(princ sPattern)(terpri)
        ; (princ "\nsNewPattern : ")(princ sNewPattern)(terpri)

        ;; Initializing Variables
        (setq lString (vl-string->list sString))
        (setq sBuildString "")
        (setq sNewString "")

        ;; Looping through each character until 
        ;; a match of the pattern is found
        (setq iItr1 -1)
        (while (> (1- (strlen sPattern)) (setq iItr1 (1+ iItr1)))
            (setq sBuildString (strcat sBuildString (chr (nth iItr1 lString))))
        );while

        (setq iItr1 (1- iItr1))
        (while (> (strlen sString) (setq iItr1 (1+ iItr1)))
            (setq sBuildString (strcat sBuildString (chr (nth iItr1 lString)))); Adding letter
            (if (> (strlen sBuildString) (strlen (vl-string-subst "" sPattern sBuildString 0)))(progn ;
                (setq sNewString (strcat sNewString (vl-string-subst sNewPattern sPattern sBuildString 0))); Updating new string
                (setq sBuildString ""); Clearing Build - Prevents double pattern remove
            ));if<-progn
        );while
        (setq sNewString (strcat sNewString sBuildString))
        (setq sReturn sNewString); New return value
    ));if<-progn

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnSubstitute - End" -1))

    sReturn ; return value
);defun fcnSubstitute
;-------------------------------------------------------------------------------
; fcnToString - Converts a value into a string.
; Function By: Garrett Beck from Hot Springs, AR, United States
; Arguments: 1
;   InputValue  = The new pattern for the string
; Syntax example: (fcnToString 1234.573) = "1234.573"
; Syntax example: (fcnToString 1234.00) = "1234"
; Syntax example: (fcnToString "  Testing Function  ") = "Testing Function"
;-------------------------------------------------------------------------------
(defun fcnToString (InputValue / sReturn OrigErrMsg)

    ;; Error Message - Start
    (if *bProgrammerDebug* (fcnBuildError "fcnToString - Start" 1))
    (setq OrigErrMsg *sErrorMessage*)
    (setq *sErrorMessage* "\nAn error occured while converting a value into a string.")

    ;; Local Function
    (defun fcnPrecisionTrim (rInputValue / sReturn sString1 sString2); Local function - Precision
        ;; Code
        (setq sString1 (itoa (fix rInputValue)))
        (setq sString2 (vl-string-right-trim "0" (rtos rInputValue 2 12)))
        (setq sReturn  (if (> (+ (strlen sString1) 2)(strlen sString2)) sString1 sString2))
        sReturn
    );defun fcnPrecisionTrim

    ;; String variable
    (setq sReturn (cond ; Return value
        ((= (type InputValue) 'INT) (itoa InputValue))
        ((= (type InputValue) 'REAL)(fcnPrecisionTrim InputValue))
        ((= (type InputValue) 'STR) (vl-string-trim " " InputValue))
        (T ""); Else statement
    ));setq<-cond
    (if *bProgrammerDebug* (setq *sErrorMessage* (strcat *sErrorMessage* "\nsReturn : " sReturn)))

    ;; Error Message - End
    (setq *sErrorMessage* OrigErrMsg)
    (if *bProgrammerDebug* (fcnBuildError "fcnToString - End" -1))
    
    sReturn
);defun fcnToString
;-------------------------------------------------------------------------------
; fcnClearLogs - Clears ATIN logs in the Temp folders after 2 days in a month.
;                Between months or years, it increases up to 3 or 4 days due to
;                uncalculated certainty.
; Function By: Garrett Beck from Hot Springs, AR, United States
; Arguments: 0
;-------------------------------------------------------------------------------
(defun fcnClearLogs (/ 
    sRawDateAndTime lDocLogs sDoc
    iNowYear iNowMonth iNowDay iDocYear iDocMonth iDocDay
    ); Local Declarations

    ;; Current Date
    (setq sRawDateAndTime (rtos (getvar "CDATE") 2 6))
    (setq iNowYear  (atoi (substr sRawDateAndTime 1 4)))
    (setq iNowMonth (atoi (substr sRawDateAndTime 5 2)))
    (setq iNowDay   (atoi (substr sRawDateAndTime 7 2)))
    
    ;; List of relevant logs
    (setq lDocLogs (vl-directory-files (fcnTemporaryFolderLocation) "CADAttInLog - *.csv"))
    (foreach sDoc lDocLogs

        ;; Document Date
        (setq iDocYear  (atoi (substr (vl-filename-base sDoc) 14 4)))
        (setq iDocMonth (atoi (substr (vl-filename-base sDoc) 19 2)))
        (setq iDocDay   (atoi (substr (vl-filename-base sDoc) 22 2)))

        ;; Retension Time (2 day minimum)
        (if (and (= iNowYear iDocYear)(= iNowMonth iDocMonth)(<= iNowDay (+ 2 iDocDay)))
            (setq sDoc nil)
        (if (and (= iNowYear iDocYear)(<= iNowMonth (1+ iDocMonth)(<= iNowDay 2)))
            (setq sDoc nil)
        (if (and (<= iNowYear (1+ iDocYear))(= iNowMonth 1)(<= iNowDay 2))
            (setq sDoc nil)
        )));if

        ;; Clearing Expired Documents
        (if sDoc (vl-file-delete (strcat (fcnTemporaryFolderLocation) "\\" sDoc)))
    );foreach
    (princ)
);defun fcnClearLogs

;;; Other Functions ====================================================== Other Functions ;;;
;;; ====================================================================================== ;;;
;;; Error Handling ======================================================== Error Handling ;;;

(defun fcnResetError ()
    (setq *iDepthErr* 0)
    (setq *iIterErr* 0)
    (setq *sDebugMessage* "")
    (setq *sErrorMessage* "")
);defun fcnResetError

(defun fcnBuildError (sMessage iAddSubtractDepth / iDepth)

    ;; Depth Counter
    (setq iDepth (cond 
        ((< 0 iAddSubtractDepth) 1)
        ((> 0 iAddSubtractDepth) -1)
        (T 0)
    ));setq<-cond
    
    ;; Error Message
    (if *bProgrammerDebug* (progn 
        
        ;; Line Number
        (setq *sDebugMessage* (strcat *sDebugMessage* "\n"  (rtos (setq *iIterErr* (1+ *iIterErr*)) 2 0) ") "))
        
        ;; Adding Character(s) for depth
        (if (> iDepth 0)(setq *iDepthErr* (1+ *iDepthErr*)))
        (repeat *iDepthErr* (setq *sDebugMessage* (strcat *sDebugMessage* "- ")))
        (if (< iDepth 0)(setq *iDepthErr* (1- *iDepthErr*)))
        (if (< *iDepthErr* 0)(setq *iDepthErr* 0))
        
        ;; Adding description
        (setq *sDebugMessage* (strcat *sDebugMessage* sMessage))
    ));if
);defun fcnBuildError

(defun fcnErrorFunction (msg / bError objFile)
    
    ;(if osm (setvar 'osmode osm)) ; Updates system variables (if applicable)
    (C:ATPort-Release-Apps T)
    (setq bError (not (member msg '("Function cancelled" "quit / exit abort"))))

    ;; Print Messages
    (if bError (progn
        (if *bProgrammerDebug* (progn   
            (princ *sDebugMessage*)(terpri);------------------------; General Path
            (princ *sErrorMessage*)(terpri);------------------------; Detailed Error
        ));if<-progn
        (princ msg)(terpri);----------------------------------------; Default Error Message
        (if (not *bProgrammerDebug*)(princ *sErrorMessage*))(terpri); User message

        ;; Save error
        (setq objFile (open *sErrDumpPath* "w"))
        (write-line (strcat "*sDebugMessage*: " *sDebugMessage*) objFile)
        (write-line "\n---------------------------\n" objFile)
        (write-line (strcat "*sErrorMessage*: " *sErrorMessage*) objFile)
        (write-line "\n---------------------------\n" objFile)
        (write-line (strcat "msg: " msg) objFile)
        (close objFile)
    ));if<-progn

    ;; Cleaning up items
    (setq *error* *OriginErrState*)
    (princ)
);defun fcnErrorFunction

;;; Error Handling ======================================================== Error Handling ;;;
;;; ====================================================================================== ;;;
;;; Function Use Global Variables ========================== Function Use Global Variables ;;;

(setq *ExternalFileType* (fcnLoadDefault "C:ATPort-With"))
(setq *sErrDumpPath* (strcat (fcnTemporaryFolderLocation) "\\AttErrDumpFile.log"))
(fcnClearLogs)

;;; Function Use Global Variables ========================== Function Use Global Variables ;;;
;;; ====================================================================================== ;;;
;;; Messages ==================================================================== Messages ;;;

(princ (strcat "Additional attribute Import and Export options have been added. Type"
            " ATHELP to see the list of added commands from this lsp file."
            "\nATOUT is set to export into " *ExternalFileType* " files. This can be"
            " changed with ATPORT-WITH.\n"))

;;; Messages ==================================================================== Messages ;;;
;;; ====================================================================================== ;;;
