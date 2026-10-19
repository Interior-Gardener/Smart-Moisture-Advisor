# SMA major project deliverables

All final files are in this folder. Nothing has been committed or published.

## Final files

- `C:\Users\karti\Desktop\All Desktop Apps\sma\presentation\LY Major Project\SMA_Major_Project_Report_10_Pages.docx`
- `C:\Users\karti\Desktop\All Desktop Apps\sma\presentation\LY Major Project\SMA_Major_Project_Report_10_Pages.pdf`
- `C:\Users\karti\Desktop\All Desktop Apps\sma\presentation\LY Major Project\SMA_Major_Project_Presentation.pptx`
- `C:\Users\karti\Desktop\All Desktop Apps\sma\presentation\LY Major Project\SMA_Major_Project_Presentation.pdf`
- `C:\Users\karti\Desktop\All Desktop Apps\sma\presentation\LY Major Project\SMA_Major_Project_Demo.mp4`

## Validation

- Report: exactly 10 pages confirmed in Word and in the exported PDF.
- Allocation: introduction on page 1; literature and gaps on pages 2-3; problem, objectives, and scope on page 4; architecture and implementation on pages 5-6; outputs and evaluation on pages 7-8; conclusion and future scope on page 9; references on page 10.
- Every report page and all 11 presentation slides were rendered and visually checked.
- Presentation preserves the supplied 13.3333 x 7.5 inch template, theme, masters, Somaiya branding, red borders, and Times New Roman typography.
- Editable native text and tables are retained in the presentation. Diagrams and screenshots are embedded images.
- Guide: Dr. Nilesh Yadav. Students: Kartik Verma and Kushal Soni.
- Old-project text, images, references, and attributions were replaced. No presentation video section, video link, video embed, or QR code remains.
- Demo: 148.17 seconds (2 minutes 28 seconds), 1920 x 1080, H.264, captioned without an audio narration track. All 889 frames decoded successfully.
- Demo records real Chrome interactions with the unchanged application. Browser capture reported no application errors.
- All 23 existing application tests passed. Software tests do not establish field performance.
- Results clearly distinguish agronomy-generated targets and scenario estimates from measured field outcomes.

## Files preserved in dump

- `LY_Major Project (1).pptx`: original unrelated-project presentation moved from the parent presentation folder after checking the backup SHA-256 hash.
- `Original_LY_Major_Project_Template.pptx`: byte-identical backup retained before authoring.
- `revision_1/`: superseded generated report and presentation drafts and their PDF exports, plus the original MPEG-4 video encode.
- `template_preview/`: original template slide renders.
- `build/`: reproducible authoring sources, actual browser captures, recording frames, assets, QA images, tool dependencies, and validation records.

Existing files in the repository report folder and the older SMA presentation in the parent presentation folder were left in place. No source code, models, dataset, or runtime dependencies were moved. No files were deleted.

## Remaining items

No deliverable is blocked. Voice narration is optional; the finished demo uses captions. Field trials and live-data integrations remain future project work, as documented in the report.

## Reproduction notes

Authoring and rendering sources are retained in dump/build. Native DrawingML editing preserved the PowerPoint template because the Artifact Tool runtime and dependency-loader tool were not exposed in this Windows session. The packaged DOCX renderer could not start because pdf2image was unavailable, so the report was rendered and page-counted through installed Word instead. The PowerShell Word COM binder stalled on export; export_word.vbs completed native PDF export successfully. These workarounds do not affect the delivered files.

Detailed hashes, page allocation, video checks, and source provenance are in dump/build/final_validation.json and dump/build/image_provenance.json.
