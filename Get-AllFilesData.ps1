<#
.SYNOPSIS
    Scans selected files or folders to extract and display detailed metadata.

.DESCRIPTION
    This script provides a GUI to select files or folders for a deep metadata scan. It uses two
    primary methods for data extraction:
    1. For image files (like JPG, PNG, TIFF), it reads the raw EXIF (Exchangeable image file
       format) data, providing detailed information about the camera, settings, and location if
       available.
    2. For all other files, it uses the Windows Shell COM object to retrieve a comprehensive list
       of metadata properties, similar to the details you would see in the File Explorer
       properties pane.

    The script features a sophisticated UI to filter by file extension, a "Find Only" mode to
    quickly list files without processing, and the ability to export all gathered data to a CSV
    file.

.EXAMPLE
    PS C:\> .\Get-AllFilesData.ps1

    Launches the main GUI window. From here, you can select files or folders to scan and view the
    extracted metadata in a new results window.

.NOTES
    Name:           Get-AllFilesData.ps1
    Version:        1.0.0, 10/18/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET Framework access for Windows Forms and COM object
                    interaction.
#>
Clear-Host
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$ExternalButtonName = "Get File(s) Data"
$ScriptDescription = "Scans selected files or folders to extract and display detailed metadata, including EXIF data for images and technical properties for other files."

# The data from "FileExtensions.csv" is now hardcoded into the script.
# This removes the dependency on the external CSV file.
# Please populate the list below with the full content of your CSV file.
$csvContent = @"
Ext,Description,Used by
*,All Files,
3DS,Cartridge game format for Nintendo 3DS,Nintendo 3DS Family
3DSX,Homebrew file for Nintendo 3DS,Homebrew launcher (3ds)
3G2,Mobile phone video,
3GP,Mobile phone video,
3GX,Luma3DS Plugin,Luma3DS
3MF,3D manufacturing format,3D builder
7z,"A compressed archive file format that supports several different data compression, encryption and pre-processing algorithms",7-zip
4TH,Forth language source code file,Forth development systems
A,Archive file,ar (Unix)
AAC,Advanced Audio Coding file,"iOS, Nintendo DSi, Nintendo 3DS, YouTube Music"
ACCDB,Microsoft Access Database,Microsoft Access Database (Open XML)
ACCFT,Microsoft Access Data Type Template,Microsoft Access
ACO,Adobe color palette format (.aco),Adobe Photoshop
ADT,"Abrechnungsdatentransfer, an xDT application",Healthcare providers in Germany
ADX,Document,Archetype Designer
ADZ,Amiga Disk Zipped (See Amiga Disk File),GZip
AGDA,Agda (programming language) source file,Agda typechecker/compiler
AGR,ArcView ASCII grid,
AHK,AutoHotkey script file,AutoHotkey
AI,Adobe Illustrator Artwork,Adobe Illustrator
AIFF,Audio Interchange File Format,professional audio processing applications and on Macintosh
AIFC,Compressed Audio Interchange File Format,
AIO,APL programming language file transfer format file,
AMF,Additive Manufacturing File Format,Computer Aided Design Software
AMG,System image file,ACTOR
AML,AutomationML,AutomationML Group
AMLX,Compressed and packed AutomationML file,AutomationML Group
AMPL,AMPL source code file,AMPL
AMR,Adaptive Multi-Rate audio,
AMV,Actions Media Video,
ANI,Animation cursors for Win,Win95 - WinNT
ANN,Annotations of old Windows Help file,Windows 3.0 - XP
APE,Monkey's Audio (Lossless),audio media players
APK,Android application package,Android
APK,Alpine Linux Package,Alpine Linux and derivatives
ARC,ARC (file format),
ART,Gerber format,"Cadence Allegro, EAGLE"
ASAX,ASP.NET global application file,
ASCX,ASP.NET User Control,
ASF,Advanced Streaming Format (Compressed Windows audio/video),Microsoft Corporation
ASHX,ASP.NET handler file,
ASM,Assembler language source,"TASM, MASM, NASM, FASM"
ASPX,Active server page extended file,Microsoft Corporation
ASX,"Advanced Stream Redirector file, redirects to an ASF file (see ASF)",Microsoft Corporation
ATG,Coco/R LL(1) formal grammar,Coco/R
AT3,Atrac 3 Sound/music file,All Sony devices and programs with the Atrac 3+ specification
AU,audio file,
AVI,Audio Video Interleave,Video for Windows
AVIF,AV1 Image File Format,
AWK,AWK script/program,"awk, GNU Awk, mawk, nawk, MKS AWK, Awka (compiler)"
AX,DirectShow Filter,Microsoft Corporation (Video Players)
AXF,lightweight geodatabase,ESRI ArcPad
B,BASIC language source,
B,bc arbitrary precision calculator language file,Unix bc tool
B64,base64 binary-to-text encoding,
BAK,backup,various
BAR,Broker Archive. Compressed file containing number of other files for deployment.,IBM App Connect
BAS,BASIC language source,QuickBASIC - GW-BASIC - FreeBASIC - others
BAT,Batch file,"MS-DOS, RT-11, DOS-based command processors"
BDF,"Glyph Bitmap Distribution Format, a format used to store bitmap fonts.",Adobe
BDT,"Behandlungsdatentransfer, an xDT application",Healthcare providers in Germany
BEAM,Executable bytecode file in fat binary format,BEAM (Erlang virtual machine)
BIB,Bibliography database,BibTex
BIN,binary file,Every OS
BLEND,Blender project file,Blender
BM3,UIQ3 Phone backup,
BMP,OS/2 or Win graphics format (BitMap Picture),QPeg - CorelDraw - PC Paintbrush - many
BPS,WPS backup file,"Microsoft Word, Microsoft Works"
BR7,Bryce 7 project file,Bryce 7
BSK,Bryce 7 sky presets file,Bryce 7
BSON,JSON-like binary serialization,MongoDB
BSP,Binary space partitioning tree file,Quake-based game engines
BYU,3D geometry format,CAD systems
BZ2,Archive,bzip2
C--,C-- language source,Sphinx C--
C,C language source,"Watcom C/C++, Borland C/C++, gcc and other C compilers"
C,Unix file archive,COMPACT
C++,C++ language source,
CPP,C++ language source,
C32,COMBOOT Executable (32-bit),SYSLINUX
CAB,Cabinet archive,"Windows 95 and later, many file archivers"
CBL,COBOL language source,
CBT,COMBOOT Executable (incompatible with DOS COM files),SYSLINUX
CC,C++ language source,
CD,ASP.NET class diagram file,
CDF,Common Data Format,
CDF,Computable Document Format,Mathematica
CDP,Trainz Railroad Simulator Content Dispatcher Pack,Trainz Railroad Simulator
CDR,Vector graphics format (drawinF,CorelDraw
CDXML,"MIME type: chemical/x-cdxmlXML version of the ChemDraw Exchange format, CDX.",
CER,Security certificate,Microsoft Windows
CGM,Computer Graphics Metafile vector graphics,A&L - HG - many
CHM,Compiled Help File,"Microsoft Windows, Help Explorer Viewer"
CHO,ChordPro lead sheet (lyrics and chords),ChordPro and similar tools
CIA,Decrypted Nintendo 3DS ROM cartridge,Nintendo 3DS
CIF,Crystallographic Information File,"RasMol, Jmol"
CLASS,Java class file,Java
CLS,ooRexx class file,ooRexx
CMD,Command Prompt batch file,Microsoft Windows NT based operating systems
CMD,executable programs,CP/M-86 operating system
CML,"Chemical Markup Language, for interchange of chemical information.",
CMOD,Celestia Model,Celestia
CN1,CNR IDL,MITRE
CNOFF,"3D object file format with normals (.noff, .cnoff) NOFF is an acronym derived from Object File Format. Occasionally called CNOFF if color information is present.",
COB,COBOL language source,GnuCOBOL
COE,Coefficient file,Xilinx ISE
COFF,"3D object file format (.off, .coff) OFF is an acronym for Object File Format. Used for storing and exchanging 3D models. Occasionally called COFF if color information is present.",
COL,DIMACS graph data format.,
COM,DOS program,DOS-
COMPILE,ASP.NET precompiled stub file,
CONFIG,Configuration file,
CPC,Compressed image,Cartesian Perceptual Compression
CPIO,cpio archive file,cpio
CPL,Control panel file,Windows 3.x
CPY,COBOL source copybook file,
CR2,Raw image format,Canon digital cameras
CR3,Raw image format,Canon R series cameras
CRAI,CRAM index,
CRAFT,Holds Spacecraft Assembly information,Kerbal Space Program
CRT,Security certificate,Microsoft Windows
CS,C# language source,
CSPROJ,C# project file,Microsoft Visual Studio
CSS,Cascading style sheet,
CSO,"Compiled Shader Object, extension of compiled HLSL",High-Level Shading Language
CSV,Comma Separated Values text file format (ASCII),
CUB,Used by electronic structure programs to store orbital or density values on a three-dimensional grid.,
CUBE,same as .cub,
CUR,Non-animated cursor (extended from ICO),Windows
D,D Programming Language source file,DMD
D,Directory containing configuration files (informal standard),Unix
DAA,Direct Access Archive,
DAE,COLLADA file,
DAF,Data file,Digital Anchor
DPFHTML,Extended from HTML,DarkPurpleOF's Website Extension
DART,Dart (programming language) source file,
DAT,AMPL data file,AMPL
DAT,"LDraw (Sub)Part File, 3D Model",LDraw
DAT,Data,RSNetWorx Project
DAT,Data file in special format or ASCII,
DAT,Database file,Clarion (programming language)
DAT,"Norton Utilities disc image data. It saves Boot sector, part of FAT and root directory in image.DAT on same drive.",Norton Utilities
DAT,"Optical disc image (can be ISO9660, but not restricted to)","cdrdao, burnatonce"
DAT,Video CD MPEG stream,
DAT,"Windows registry hive (REG.DAT Windows 3.11; USER.DAT and SYSTEM.DAT Windows 95, 98, and ME; NTUSER.DAT Windows NT/2000/XP/7)",Microsoft Windows
DATS,Dynamic source,ATS
DB,Database file,DB Browser for SQLite
DBA,DarkBasic source code,
DBC,Database Connection configuration file,AbInitio
DBF,Native format of the dBASE database management application.,
DBG,Debugger script,DOS debug - Watcom debugger
DBG,Symbolic debugging information,Microsoft C/C++
DEB,deb software package,Debian Linux and derivatives
DEM,digital elevation model (DEM) including GTOPO30 and USGSDEM. GTOPO30 is a distribution format for a global digital elevation model (DEM) with 30-arc-second grid spacing. USGSDEM is the standard format for the distribution of terrain elevation data for the United States.,
DGN,CAD Drawing,"Bentley Systems, MicroStation and Intergraph's Interactive Graphics Design System (IGDS) CAD programs"
DICOM,Digital Imaging and Communications in Medicine (DICOM) bitmap,DICOM Software (XnView)
DIF,Data Interchange Format,Visicalc
DIF,Output from [diff] command - script for Patch command,
DIRED,Directory listing (ls format),Dired
DIVX,DivX media format,
DMG,Apple Disk Image,macOS (Disk Utility)
DMP,memory dump file (e.g. screen or memory),
DN,Dimension model format,Adobe Dimension
DNG,"Digital Negative, a-publicly available archival format for the raw files generated by digital cameras","At least 30 camera models from at least 10 manufacturers, and at least 200 software products"
DOC,"A Document, or an ASCII text file with text formatting codes in with the text; used by many word processors",Microsoft Word and others
DOCM,Microsoft Word Macro-Enabled Document,Microsoft Word
DOCX,Microsoft Word Document,Microsoft Word
DOT,Microsoft Word document template,Microsoft Word
DOTX,Office Open XML Text document template,Microsoft Word
DPX,Digital Picture Exchange,
DRC,Dirac format video,
DSC,Celestia Deep Space Catalog file,Celestia
DTA,Stata database transport format.,Stata
DTD,Document Type Definition,
DVC,Data version control yaml pointer into blob storage,
DWF,Autodesk Design Web Format,Design Review
DWG,Drawing,"AutoCAD, IntelliCAD, PowerCAD, Drafix, DraftSight etc."
DX,same as JDX and JCM.,
DXF,Drawing Interchange File Format vector graphics,"AutoCAD, IntelliCAD, PowerCAD, etc."
E,E language source code,E
E##,EnCase Evidence File chunk,EnCase Forensic Analysis Suite entity
E00,ArcInfo interchange file,GIS software
E2D,2-dimensional vector graphics file,Editor included in JFire
e57,A file format developed by ASTM International for storing point clouds and images,Most software that enables viewing and/or editing of 3D point clouds
EBD,"versions of DOS system files (AUTOEXEC.BAT, COMMAND.COM, CONFIG.SYS, WINBOOT.SYS, etc.) for an emergency boot disk","Windows 98, ME"
EC,Source code,eC
ECC,Error-checking file,dvdisaster
EDE,Ensoniq EPS disk image,AWAVE
EDF,European data format,Medical timeseries storage files
EFI,Extensible Firmware Interface,
EIS,EIS Spectrum Analyser Project,EIS Spectrum Analyser Archived 2010-03-29 at the Wayback Machine
EL,Emacs Lisp source code file,Emacs
ELC,Byte-compiled Emacs Lisp code,Emacs
ELF,"Executable and Linkable File, EurekaLog
File (contains details of an exception)","Unix
EurekaLog (https://www.eurekalog.com/)"
EMAIL,Outlook Express Email Message,"Windows Live Mail, Outlook Express, Microsoft Notepad"
EMAKER,E language source code (maker),E
EMF,Microsoft Enhanced Metafile,
EML,Email conforming to RFC 5322; Stationery Template,Email clients;Outlook Express
EMZ,Microsoft Enhanced Metafile compressed with ZIP,Microsoft Office suite
EOT,Embedded OpenType,
EP,GUI wireframe/prototype project,"Prikhi Pencil, Evolus Pencil"
EPA,Award BIOS splash screen,"Award BIOS, XnView"
EPS,Encapsulated PostScript,CorelDraw - PhotoStyler - PMView - Adobe Illustrator - Ventua Publisher
EPUB,Electronic Publication (e-Reader format),Okular (Linux) - Apple Books - Sony Reader - Adobe Digital Editions - Calibre (LMW)
EU4,Europa Universalis 4 save game file,Europa Universalis 4
ERL,Erlang source code file,
ES6,ECMAScript 6 file,
ESCPCB,"Data file of ""esCAD pcb"", PCB Pattern Layout Design Software",esCAD pcb provided by Electro-System
ESCSCH,"Data file of ""esCAD sch"", Drawing Schematics Diagram Software",esCAD sch provided by Electro-System
ESD,Windows Imaging Format,"ImageX, DISM, 7-Zip, wimlib"
ETL,event trace log file,Microsoft
EVT,Windows Event log file,Microsoft Windows NT 4.0 - XP; Microsoft Event Viewer
EVTX,Windows Event log file XML structured,"Microsoft Windows Vista, 7, 8; Microsoft Event Viewer"
EX,Elixir source code file,Elixir programming language running on BEAM (Erlang virtual machine)
EXE,Directly executable program,"DOS, OpenVMS, Microsoft Windows, Symbian or OS/2"
EXP,Drawing File format,Drawing Express
EXP,Melco Embroidery Format,Embroidermodder
EXR,"OpenEXR raster image format (.exr). Used in digital image manipulation for theatrical film production. EXR is an acronym for Extended Dynamic Range. Stores 16 bit per pixel IEEE HALF-precision floating-point color channels. Can optionally store 32-bit IEEE floating-point ""Z"" channel depth-buffer components, surface normal directions, or motion vectors.",
EXS,Elixir script file,Interactive Elixir (IEx) shell
F,Forth language source code file,Forth development systems
F,Fortran language source code file (in fixed form),Many Fortran compilers
F01,Fax,perfectfax
F03,Fortran language source code file (in free form),Many Fortran compilers
F08,Fortran language source code file (in free form),Fortran compilers
F18,Fortran language source code file (in free form),Fortran compilers
F4,Fortran IV source code file,Fortran IV source code
F4V,A container format for Flash Video that differs from the older FLV file format (see also SWF),Adobe Flash
F77,Fortran language source code file (in fixed form),Many Fortran compilers
F90,Fortran language source code file (in free form),Many Fortran compilers
F95,Fortran language source code file (in free form),Many Fortran compilers
FA,FASTA format sequence file,
FAA,FASTA format amino acid,
FACTOR,Factor source file,
FASTA,FASTA format sequence file,
FASTQ,FASTQ format sequence file,
FB,Forth language block file,Forth development systems
FB2,FictionBook e-book 2.0 file (DRM-free XML format),e-book readers
FBX ,"3D model geometry, material textures, lighting, armature, and animation sequences for inter application use/transport",Autodesk
FCHK,Gaussian-formatted checkpoint file. Used to store result of quantum chemistry calculation results.,
FCS,FCS molecular biology format.,
FEN,Forsyth–Edwards Notation,Chess applications -text data describes a specific position.
FF,Farbfeld image,
FFAPX,FrontFace plugin file,FrontFace digital signage software
FFN,FASTA format nucleotide of gene regions,
Fit,CurvFit Input file format,CurvFit software
FITS,Flexible Image Transport System,astronomy software
FLAC,"Audio codec, Audio file format",
FLAME,Fractal configuration file,Apophysis
FLP,FL Studio project file,FL Studio
FLV,A container format for Flash Video (see also SWF),Adobe Flash
FMU,A Functional Mockup Unit (FMU) implements the Functional Mockup Interface (FMI).,
FNA,FASTA format nucleic acid,
FNI,FileNet Native Document,FileNet
FNX,Saved notes with formatting in markup language,FeatherNote
FODG,OpenDocument Flat XML Drawings and vector graphics,"OpenDocument, LibreOffice, Collabora Online"
FODP,OpenDocument Flat XML Presentations,"OpenDocument, LibreOffice, Collabora Online"
FODS,OpenDocument Flat XML Spreadsheets,"OpenDocument, LibreOffice, Collabora Online"
FODT,OpenDocument Flat XML Text Document,"OpenDocument, LibreOffice, Collabora Online"
FOR,Fortran language source file,many
FQL,Fauna Query Language source file,Fauna
Freq,Match-n-Freq (TM) Input file format,Match-n-Freq software
FRAG,"Fragment File, usually stored on MOVPKG files",MOVPKG
FRM,MySQL Database Metadata,MySQL
Frq7,Match-n-Freq (TM) version 7+ Input file format,Match-n-Freq software
FS,F# source file,F# compilers
FS,Forth language source code file (often used to distinguish from .fb),Forth development systems
FTH,Forth language source code file,Forth development systems
G4,ANTLR4 grammar,ANTLR4
G6,Graph6 graph data format. Used for storing undirected graphs.,
GB,GenBank molecular biology format.,
GBK,same as GB,
GBR,Gerber format,PCB CAD/CAM
GDF,General Data Format for Biomedical Signals,"Biomedical signal processing, Brain Computer Interfaces"
GDSCRIPT,Godot (game engine) Script File,Godot (game engine)
GDT,"Gerätedatentransfer, an xDT application",Healthcare providers in Germany
GED,GEDCOM,Genealogy data exchange
GEOJSON,GeoJSON is specified by RFC 7946.,
GEOTIFF,Used for archiving and exchanging aerial photography or terrain data.,
GGB,GeoGebra File,GeoGebra
GIF,Compuserves' Graphics Interchange Format (bitmapped graphics),QPeg - Display - CompuShow
GM9,Scripts for GodMode9,Godmode9
GMI,Gemtext markup language,Gemini (Protocol)
GMK,GameMaker Project File,YoYo Games
GML,Geography Markup Language File,Geography Markup Language
GML,Used for the storage and exchange of graphs. Native format of the Graphlet graph editor software.,
GO,Go source code file,Go (Programming language)
GODOT,Godot (game engine) Project File,Godot (game engine)
GPX,GPS eXchange Format,
GRAPHML,"GraphML is an acronym derived from Graph Markup Language. Represents typed, attributed, directed, and undirected graphs.",
GRB,Commonly used in meteorology to store historical and forecast weather data. Represents numerical weather prediction output (NWP).,
GREXLI,Uncompressed folder as a file,WinGrex/GrexliLib
GRIB,same as GRB,
GRP,Data pack files for the Build Engine,Duke Nukem 3D
GSF,Generic sensor format,"Used for storing bathymetry data, such as that gathered by a multibeam echosounder."
GTF,Gene transfer format,
GV,Graph Visualization,Graphviz
GW,"same as .lgr. See LGR below for details on LEDA (.gw, .lgr).",
GXL,"GXL is an acronym derived from Graph Exchange Language. Represents typed, attributed, directed, and undirected graphs.",
GZ,gzip compressed data,gzip
H!,On-line help file,Flambeaux Help! Display Engine
H!,Pertext database,HELP.EXE
H--,C-- language header,Sphinx C--
H,Header file (usually C language),Watcom C/C++
H++,Header file,C++
HA,Archive,HA
HACK,Source file for the programming language hack,The HHVM
HAR,HTTP Archive format (JSON-format web-browser log),W3C draft standard
HDF,General-purpose format for representing multidimensional datasets. Developed by the US National Center for Supercomputing Applications (NCSA).,
HDF5,General-purpose format for representing multidimensional datasets and images. Incompatible with HDF Version 4 and earlier.,
HDI,Hard Disk Image file (PC-9800 disk image file),PC-9800 emulators
HDMP,heap dumpfile,
HEIC,HEIF raster image and compression format. Commonly used for storing still or animated images.,
HEIF,same as HEIC.,
HH,C++ header file,
HIN,"HyperChem HIN format. Used in cheminformatics applications and on the web for storing and exchanging 3D molecule models. Maintained by HyperCube, Inc.",
HOF,Basic Configuration file,OMSI The Bus Simulator
HOI4,Hearts Of Iron 4 file,Hearts Of Iron 4 Save game
HPP,C++ header file,Zortech C++ - Watcom C/C++
HTA,HTML Application,Microsoft Windows
HTM,see HTML,
HTML,Hypertext Markup Language (WWW),Netscape - Mosaic - many
HUM,3D Model database,OMSI The Bus Simulator
HXX,C++ header file,
ICAL,same as ICS (see below).,
ICC,ICC profile,Color Configuration for color input or output devices
ICE,LHA Archive,"Cracked LHA (old LHA), Total Commander"
ICL,Icon library,Microsoft Windows
ICNS,Macintosh icons format. Raster image file format.,
ICO,Icon file,ICONEDIT.EXE; Microsoft Windows
ICS,ICS iCalendar format. Used for the storage and exchange of calendar information. Commonly used in personal information management systems.,
IFB,Same as ICS (see above).,
IFC,"Industry Foundation Classes (platform neutral, open file format used by BIM software). It is registered by ISO and is an official International Standard ISO 16739-1:2018.",BIM software
IGC,Flight tracks downloaded from GPS devices in the International Gliding Commission's prescribed format,
IGES,Initial Graphics Exchange Specification,
IMG,Disk image,
INFO,Texinfo hypertext document,"Texinfo, info (Unix), Emacs"
INI,Configuration file,
IO,Archive,CPIO
IPT,XnView IPTC template,"XnView, XnViewMP"
IPTRACE,AIX iptrace captures dualhome.iptrace (AIX iptrace) Shows Ethernet and Token Ring packets captured in the same file.,WireShark
IQBLOCKS,Main file for programming a VEX Robot,VEX Robotics
IRX,"""IOP Relocatable eXecutable"". Library files to dynamically link application code to the Input/Output Processor on the PS2 to communicate with devices like memory cards, USB devices, etc.",Sony PlayStation 2
ISO,ISO-9660 table,see: List of ISO image software
IT,Impulse Tracker music file,Impulse Tracker
JL,Julia script file,Julia (programming language)
J2C,JPEG 2000 image,JPEG 2000
J2K,JPEG2000 raster image and compression format. Can store images as an array of rectangular tiles that are encoded separately.,
JAR,Java archive,"JAR, Java Games and Applications"
JAV,see JAVA,
JAVA,Java source code file,
JBIG,Joint Bilevel Image Group,
JCM,same as JDX (see below).,
JDX,Chemical spectroscopy format. JCAMP is an acronym derived from Joint Committee on Atomic and Molecular Physical Data.,
JNLP,Java Network Launching Protocol,Java Web Start
JP2,JPEG 2000 image,
JPE,Joint Photographic Experts Group graphics file format,Minolta/Konica Minolta cameras use this for JPEGs in Adobe RGB color space
JPEG,Joint Photographic Experts Group graphics file format,QPeg - FullView - Display
JPG,Joint Photographic Group,various (Minolta/Konica Minolta cameras use this for JPEGs in sRGB color space)
JS,JavaScript file,script in HTML pages
JSON,JSON (JavaScript Object Notation),Ajax
JSP,Jakarta Server Pages,Dynamic pages running Web servers using Java technology
JUMP,Beyond Jump save file,Cell to Singularity
JVX,JavaView 3D geometry format. The native format of the JavaView visualization software. Used for the visualization of 2D or 3D geometries. Can be embedded in web pages and viewed with the JavaView applet.,
JXL,JPEG XL raster graphics file,
KEY,Keynote Presentation,Keynote
KLC,MSKLC Source file,Microsoft Keyboard Layout Creator
KML,Keyhole Markup Language,Google Earth
KMZ,Keyhole Markup Language (Zip compressed),Google Earth
KO,Linux kernel module format system file,Linux
KRA,Krita image file,Krita
KRABER,Kraber source code file,Kraber Programming Language
KSH,Kornshell source file,Kornshell
KT,Kotlin source code file,Kotlin Programming Language
KV,Kivy,
LABEL,Dymo label file,Dymo desktop software
LATEX,LaTeX typesetting system and programming language. Commonly used for typesetting mathematical and scientific publications.,
LBR,.LBR Archive,for CP/M and MS-DOS using the LU program
LDB,.LDB Leveldb data file,Google key-value storage library
LDB,.LDB MDB Database lock file,Microsoft Access Database
LDT,"Labordatenträger, an xDT application",Healthcare providers in Germany
LGR,"LEDA graph data format. Commonly used exchange format for graphs. Stores a single, typed, directed, or undirected graph. Native graph file format of the LEDA graph library and the GraphWin application. LEDA is an acronym for Library of Efficient Datatypes and Algorithms.",
LHA,LHA Archive,"LHA/LHARC, Total Commander"
LISP,LISP source code file,
LL,LLVM Assembly Language,llvm
LM,Language Model File,Microsoft Windows
LMD,FCS molecular biology format.,
LNK,Local file shortcut,Microsoft Windows
LOGICX,Logic Pro files,"Logic Pro (MacOS, iPadOS)"
LRC,Lyrics file,Lyrics for karaoke-related system and program
LUA,Lua script file,Lua (programming language)
LWO,Native format of the LightWave 3D rendering and animation software. LWO is an acronym for LightWave Object. Developed by NewTek. Stores 3D objects as a collection of polygons and their properties.,
LZ,Archive,Lzip
M,Mathematica Package File,Mathematica
M,MATLAB M-File,MATLAB
M,Mercury Source File,Mercury
M,Source code,Objective-C
M2TS,BDAV MPEG-2 transport stream,
M3U,MPEG Audio Layer 3 Uniform Resource Locator playlist,Media players
M3U8,"MPEG Audio Layer 3 Uniform Resource Locator playlist, using UTF-8 encoding",Media players
M3U,MPEG Audio Layer 3 Uniform Resource Locator playlist,Media players
M4A,MPEG-4 Part 14 audio,iTunes Store
M4P,DRM-encumbered MPEG-4 Part 14 media,iTunes Store (formerly)
M4R,See M4A,Apple iPhone ringtones
M4V,"MPEG-4 Part 14 video, which may optionally be encumbered by FairPlay DRM","iTunes Store, Handbrake"
M64,Mupen64 gameplay recording,Mupen64
MA,"Autodesk Maya scene description format. The native format of the Maya modeling, animation, and rendering software.",
MAT,"MATLAB MAT-files. The native data format of the MATLAB numerical computation software. Stores numerical matrices, Boolean values, or strings. Also stores sparse arrays, nested structures, and more.",MATLAB and Octave
MBOX,"Unix mailbox format. Holds a collection of email messages. Native archive format of email clients such as Unix mail, Thunderbird, and many others.",
MCADDON,A zip file that contains .mcpack or .mcworld files to modify Minecraft: Bedrock Edition generally used to distribute add-ons to other users.,Minecraft: Bedrock Edition
MCMETA,A Minecraft custom resource pack configuration file.,Minecraft: Bedrock Edition
MCPACK,"A zipped resource or behavior pack that modifies Minecraft: Bedrock Edition, typically used to transfer resources between users.",Minecraft: Bedrock Edition
MCPROJECT,Minecraft Bedrock Editor's filetype. Files of this type only open in Editor and are capable of containing Editor extensions.,Minecraft: Bedrock Edition
MCSTRUCTURE,"Contains a Minecraft structure such as a building or natural feature, saved using the Structure Block tool can be shared between players, allowing the sharing of each other's structures.",Minecraft: Bedrock Edition
MCTEMPLATE,A zip archive containing the template of a world used in Minecraft.,Minecraft: Bedrock Edition
MCWORLD,"A zip archive that contains all the files needed to load a Minecraft: Bedrock Edition or Minecraft Education world, for example .dat and .txt files.",Minecraft: Bedrock Edition
MCF,Multimedia Container Format (predecessor of Matroska),
MD,Markdown-formatted text file,Markdown
MDB,MDB database file. The native format of the Microsoft Access database application. Used in conjunction with the Access relational database management system and as an exchange format.,Microsoft Access
MDF,"Master Data File, a Microsoft SQL Server file type",Microsoft SQL Server
MDF,"Measurement Data Format, a binary file format for vector measurement data","automotive industry, developed by Robert Bosch GmbH"
MDI,"Document save in high-resolution, created by MSOffice to scan documents (OCR) and turn them into a .DOC",Microsoft Office
MDG,"Digital Geometry (Programmable CAD) file format, developed by DInsight",Digital Geometric Kernel
MDL,Model,3D Design Plus or Simulink
MDMP,Mcrosoft minidump file created by Windows when a program crashes,Microsoft Windows SDK
MDS,Midi Session,Sound Imp.
MEX,MEX file (executable command),Matlab
MGF,Wolfram System MGF bitmap format. Used by the Wolfram System user interface for storing raster images. MGF is an acronym for Mathematica Graphics Format.,
MGF,Materials and Geometry Format,
MHT,MIME encapsulation of aggregate HTML documents,Web browsers
MID,Standard MIDI file,"Music synthetizers, Winamp"
MKA,Matroska audio,
MKV,Matroska video,
mlx,MATLAB live script file,MATLAB
MML,MathML mathematical markup language. Used for integrating mathematical formulas in web documents. Rendering of embedded MathML is supported by a number of browsers and browser additions.,
MNT,Surface metrology or image analysis document,MountainsMap
MO,Modelica models. File format specified by the Modelica Association.,
MOBI,eBook,"Mobipocket, Kindle"
MOD,AMPL model file,AMPL
MOD,Modula language source,
MOD,Modula-2 source code file,Clarion Modula-2
MOD,Tracker file format created for Ultimate Soundtracker for the Amiga,MOD (file format)
MOD,MOD is recording format for use in digital tapeless camcorders.,MOD and TOD
MODULES,Module,GTK+
MOE,same as Modelica Model .mo.,
MOL,MDL Molfile,RasMol
MOL2,Tripos Sybyl MOL2 Format,"SYBYL, RasMol"
MOP,MOPAC input file,"MOPAC, RasMol"
MOV,Animation format (Mac),"QuickTime, AutoCAD AutoFlix"
MP2,MPEG audio file,"Winamp, xing"
MP3,"MPEG audio stream, layer 3","AWAVE, CoolEdit(+PlugIn), Winamp, many others"
MP4,"multimedia container format, MPEG 4 Part 14",Winamp
MPA,"MPEG audio stream, layer 1,2,3",AWAVE
MPC,Musepack audio,
MPD,LDraw file (multi-part DAT file),LDraw
MPEG,"multimedia containter format, video, audio","MPEG Player, Winamp"
MPG,see MPEG,
MPS,MPS linear programming system format (.mps) Commonly used as input format by LP solvers. MPS is an acronym for Mathematical Programming System.,
MR,A Engine Simulator engine save file,Engine Simulator
MRSH,MarsShell Script File,MarsShell (mrsh)
MSC,management saved console,Microsoft; Microsoft MMC
MSAV,Mindustry map file,Mindustry
MSCH,Mindustry schematic file,Mindustry
MSDL,Manchester Scene Description Language,
MSF,Multiple sequence file (Pileup format),
MSI,Windows Installer Package,Microsoft Windows
MSO,Microsoft Outlook metadata for a Microsoft Word 2000 email attachment,Microsoft Outlook
MSSTYLES,Windows visual style file,Microsoft Windows
MSU,Microsoft Update Package,Microsoft Windows
MTS,See M2TS,
MUP,MUP -- File type used by MindMup to export editable Mind Maps,
MM,File type used by FreeMind to export editable Mind Maps,
MX,Wolfram Language serialized package format (.mx) Wolfram Language serialized package format. Used for the distribution of Wolfram Language packages. Stores arbitrary Wolfram Language expressions in a serialized format optimized for fast loading.,
MXF,"Material exchange format (RFC 4539, SMPTE 377M)",
MYD,a MyISAM data file in MySQL,"MyISAM, MySQL"
MYI,a MyISAM index file in MySQL,"MyISAM, MySQL"
NB,Wolfram Mathematica Notebook (see Wolfram Language),Mathematica
NC,Binary Data,netCDF software package
NC,Instructions for NC (Numerical Control) machine,CAMS
NC,Name code program,namec-git/namec software package
NCD,"NC Drill File (Excellon Format, printed circuit board hole definitions)",Most PCB layout software
NDK,NDK seismologic file format. Commonly used for storage and exchange of earthquake data. Stores geographical information and wave measurements for individual seismological events.,
NDS,Nintendo DS file. Used for Homebrew and official games.,Nintendo DS consoles and emulators
NEF,Nikon RAW image format,Nikon cameras
NEO,Text file; media,
NET,Pajek graph data format (.net) Pajek graph language and data format. Commonly used exchange format for graphs. The native format of the Pajek network analysis software. The format name is Slovenian for spider.,Pajek network analysis software
NEU,Pro/Engineer neutral file format,PTC Pro/Engineer
NEX,"NEXUS phylogenetic format (.nex, .nxs) Commonly used for storage and exchange of phylogenetic data. Can store DNA and protein sequences, taxa distances, alignment scores, and phylogenetic trees.",
NF,Instructions for NC machine made by TRUMPF,TRUMPF
NFO,iNFO for accompanying media files,"Kodi, Plex"
NIM,Nim source code file,Nim
NMF,Node Map File,Used by SpicyNodes
NOFF,"3D object file format with normals (.noff, .cnoff) NOFF is an acronym derived from Object File Format. Occasionally called CNOFF if color information is present.",
NPR,Nuendo Project File,Steinberg Nuendo
NRO,Nintendo Switch executable file,
NRW,Nikon Coolpix RAW image,Nikon
NRX,NetRexx Script File,NetRexx
NS1,NetStumbler file,NetStumbler
NSA,media,Nullsoft Streaming Audio
NSF,NES sound format file,Transfer of NES music data
NSV,media,Nullsoft Streaming Video
NUMBERS,Numbers spreadsheet file,Numbers
"NWD, NWF",Navisworks 3D drawing,Navisworks
NXS,"Same as NEXUS (.nex, .nxs). See NEX above for more details. NEXUS phylogenetic format (.nex, .nxs)",
O,Object file,UNIX - Atari - GCC
OBJ,Compiled machine language code,
OBJ,Object code,Intel Relocatable Object Module
OBJ,Wavefront Object,
OBS,Script,ObjectScript
OCX,OLE custom control,
ODB,Database front end document,"OpenDocument, LibreOffice"
ODF,"Formula, mathematical equations","OpenDocument, LibreOffice"
ODG,Drawings and vector graphics,"OpenDocument, LibreOffice, Collabora Online"
ODP,Presentations,"OpenDocument, LibreOffice, Collabora Online"
ODS,OpenDocument spreadsheet format (.ods),"OpenDocument, LibreOffice, Collabora Online"
ODT,Text (Word processing) documents,"OpenDocument, LibreOffice, Collabora Online"
OFF,"3D object file format (.off, .coff) OFF is an acronym for Object File Format. Used for storing and exchanging 3D models.",
OGA,Audio file in the Ogg container format,libogg
OGG,Vorbis audio in the Ogg container format,libogg
OGV,Video file in the Ogg container format,libogg
OGX,Ogg Multiplex Profile,libogg
OPUS,Ogg/Opus audio file,"mpv, mplayer, many others (official file format)."
ORG,Emacs Org mode,Emacs Org mode major mode
ORG,Older Origin Project,Origin versions 4 or earlier
OSB,Osu! Storyboard,Osu!
OSC,OpenStreetMap Changeset,OpenStreetMap
OSK,Osu! Skin,Osu!
OSM,OpenStreetMap data,OpenStreetMap
OSM,OpenStreetMap note,OpenStreetMap
OSR,Osu! Replay,Osu!
OST,Offline Storage Table,"Microsoft e-mail software: Outlook Express, Microsoft Outlook"
OSU,Osu! Beatmap Info,Osu!
OSZ,Osu! Beatmap,Osu!
OTB,Over-the-air bitmap graphics,
OTF,OpenType font,
OTL,The Vim Outliner,A vim plugin
OTF,"Formula, mathematical equations template","OpenDocument, LibreOffice, Collabora Online"
OTG,Drawings and vector graphics template,"OpenDocument, LibreOffice, Collabora Online"
OTP,Presentations template,"OpenDocument, LibreOffice, Collabora Online"
OTS,Spreadsheets template,"OpenDocument, LibreOffice, Collabora Online"
OTT,Text (Word processing) documents template,"OpenDocument, LibreOffice, Collabora Online"
OV2,Overlay file (part of program to be loaded when needed),TomTom Point of Interest
OWL,Web ontology language (OWL) file,Protégé and other ontology editors
OXT,OpenOffice.org extension,OpenOffice.org / LibreOffice
P,Database PROGRESS source code,PROGRESS
P,PASCAL source code file,
P,Parser source code file,
P8,PICO-8 Game File,PICO-8
P10,Certificate Request,
P12,Personal Information Exchange,Crypto Shell Extensions
PACK,Pack200 Packed Jar File,
PAGES,Pages document file,Pages
PAK,Archive,Pak
PAL,Paint Shop Pro color palette (JASC format),"PaintShop Pro, XnView"
PAM,PAM Portable Arbitrary Map graphics format,Netpbm
PAPA,"Flipline Studio's game backups like JackSmith, Papa's Wingeria version 1.2+, Papa's Pancakeria version 1.4+",
PAR,Parity Archive,
PAR2,Parity Archive v2,
PARAMS,"MXNet net representation format (.json, .params) Underlying format of the MXNet deep learning framework, used by the Wolfram Language. Networks saved as MXNet are stored as two separate file: a .json file specifying the network topology and a .params file specifying the numeric arrays used in the network.",
PAS,Pascal language source,Borland Pascal
PAX,pax archive file,"pax, GNU Tar"
PBLIB,Power Library,PowerBASIC
PBM,ASCII portable bitmap format (.pbm) PBM monochrome raster image format. Member of the Portable family of image formats. Related to PGM and PPM. Native format of the Netpbm graphics software package.,Netpbm graphics software package
PBO,A file type used by Bohemia Interactive,"Arma 3, PBO Manager"
PCAP,Network packet capture format (.pcap),WireShark
PCL,HP-PCL graphics data file,HP Printer Command Language
PCS,PCSurvey file,PCSurvey by Softart - land surveying software
PCX,PC Paintbrush file,PC Paintbrush
PDB,debugging data,Microsoft Windows
PDB,Molecule (protein data bank),
PDE,Processing source code file,Processing programming language
PDF,Adobe's Portable Document Format,Adobe Acrobat Reader
PDI,Portable Database Image,
PDM,Program,Deskmate
PDM,PowerDesigner's physical data model (relational model) file format,PowerDesigner
PDM,Visual Basic (VB) Project Information File,Visual Basic
PDN,Image file,Paint.NET
PDS,PALASM Design Description,
PDS,Planetary Data System Format,
PEM,A text-based certificate file defined in RFC 1421 through RFC 1424,"Applications that need to use cryptographic certificates, including web-servers"
pet,package,Puppy Linux
PFA,PostScript Font File,
PFA,Type 3 font file (unhinted PostScript font),
PFAM,PFAM format,
PFB,PostScript font,Adobe Type Manager (ATM)
PFC,"(Personal Filing Cabinet) contains e-mail, preferences and other personal information",AOL
PFM,PostScript Type 1 font metric file,"Microsoft Windows, Adobe Acrobat Reader"
PFM,Windows Type 1 font metric file,
PGN,Portable Game Notation -Text specification for Chess game,Most chess playing computer applications
PFX,An encrypted certificate file,"Applications that need to use cryptographic certificates, including web-servers"
PHF,Database File,Nuverb Systems Inc: Donarius
PHN,Phun scene,Algodoo (previously Phun)
PHP,PHP file,
PHP3,PHP 3 file,
PHP4,PHP 4 file,
PHR,Phrases,LocoScript
PHY,Phylip format,
PHZ,Algodoo scene,Algodoo
PI2,Portrait Innovations High Resolution Encrypted Image file,Portrait Innovations Studio2 (proprietary)
PIE,GlovePIE script file,GlovePIE
PIR,PIR format,
PIT,Compressed Mac file archive created by PACKIT,unpackit.zoo
PIT,Partition Information Table for Samsung's smartphone with Android,Odin3
PK3,Quake III engine game data,
PKL,Pickle file,Python object serialization
PKA,Archive,PKARC
PKG,General package for software and games,"MacOS, iOS, PSVita, PS3, PS4, PS5, Symbian, BeOS..."
PL,Perl source code file,
PL,Prolog source code file,
PL,IRAF pixel list,IRAF astronomical data format
PLI,PL/I source file,PL/I compilers
PLR,Terraria player/character file,Terraria
PLS,"Multimedia Playlist, primarily for streaming","Shoutcast, IceCast, others"
PM,Perl module,
PMA,PMarc Archive,
PMP,PenguinMod Project file,PenguinMod's project file
PNG,Portable Network Graphics file,"Web browsers, image viewing and editing applications"
POM,Build manager configuration file,Apache Maven POM file
PPEG,parsimonious PEG grammar,parsimonious parser generator
PPTX,MS Office Open-XML Presentation,Microsoft PowerPoint
PPSX,MS Office Open-XML Auto-Play Presentation,Microsoft PowerPoint
PRJ,Mkd (Unix command),Mkd project file to extract documentation
PROPERTIES,Configuration file format. Commonly used in Java projects. Associates string keys to string values.,
PROTO,Message specification,Google Protocol Buffers
PRP,Plasma Registry Page,Plasma (engine)
PS,Adobe Postscript file,PostScript
PSD,Photoshop native file format,Adobe Photoshop
PSDC,Photoshop Cloud Document,Adobe Photoshop
PSM1,Windows Powershell module,Windows Powershell
PSPPALETTE,Paint Shop Pro color palette (JASC format),Paint Shop Pro 8.0 and newer
PST,Archive File,Microsoft Outlook
PS1,Windows Powershell script,Windows PowerShell
PTF,PlayStation Portable Theme file,PSP Theme settings menu
PTF,Pro Tools Session File,Digidesign/Avid Pro Tools version 7 up to version 9
PTS,Pro Tools Session File,Digidesign Pro Tools (legacy version)
PTX,Pro Tools Session File,Avid Pro Tools version 10 or later
PUB,Public key ring file,Pretty Good Privacy RSA System
PUP,Pileup format,
PY,Python script file,Python (programming language)
QFX,Quicken-specific implementation of the OFX specification,Intuit Quicken
QIF,Quicken Interchange Format,Intuit Quicken
QLC,ATM Type 1 fonts script,Adobe Type Manager
QOI,Quite OK Image Format,"Web browsers, image viewing and editing applications"
QSS,QT Style Sheet,QT Python GUI library
QT,QuickTime movie (animation),
QTVR,QuickTime VR Movie,
R,Ratfor file,Ratfor
R,Script file,R
"R00, R01, ...",Part of a multi-file RAR archive,RAR
R2D,Reflex 2 datafile,Reflex 2
R3D,Red Raw Video (raw video data created with a Red camera),Red Camera
R8P,PCL 4 bitmap font file,Intellifont
RAD,2-op FM music,Reality AdLib Tracker
RAD,Radiance,Radiance
RAL,Remote Access Language file,Remote Access
RAM,Ramfile,RealAudio
RAP,Flowchart,RAPTOR
RAR,Archive,RAR
RAS,Graphics format,SUN Raster
RB,Ruby Script file,
RBXL,Roblox Experience file,Roblox Studio
RC,Configuration file,"emacs, Vim (text editor), Bash (Unix shell)"
RC,Resource Compiler script file,"Microsoft C/C++, Borland C++"
RDP,RDP connection,Remote Desktop connection
RDS,Data file,R
RES,Compiled resource,"Microsoft C/C++, Borland C++"
"REX, REXX",Rexx Script file,ooRexx
RKT,Racket language source file,DrRacket integrated development environment (IDE) for the Racket (programming language)
RM,RealMedia,RealPlayer
RMD ,R Markdown,RStudio
RMVB,RealMedia Variable Bitrate,RealPlayer
Rob,Robot4 (TM) Input file format,Robot4 software
ROL,AdLib Piano Roll,AdLib Visual Composer
RPM,RPM software package,"Red Hat Linux, the Linux Standard Base and several other operating systems"
RS,Rust language source,
RSM,Compressed Filetype for Mods of Mario Fangames,SMBX2
RSA,Harwell–Boeing matrix format. Used for exchanging and storing sparse matrices.,
"RSL, RSLS, RSLF",Resilio Sync File Placeholder,
RST,reStructuredText,Docutils
RTF,Rich Text Format text file (help file script),many - Microsoft Word
RUA,same as RSA,
RUN,AMPL script file,AMPL
RUN,Makeself shell self-extracting archive,shell
S,assembler source code file,Unix
S3M,Scream Tracker Module,Scream Tracker
S7I,Seed7 library / include file,Seed7 interpreter and compiler
SAIF,Spatial Archive and Interchange Format,
SASS,"Sass stylesheet language, indented-format",
SAT,ACIS ACIS .sat,
SAV,SPSS tabular data (binary),"PSPP, SPSS"
SB,Scratch 1.x project,Scratch
SB2,Scratch 2.0 project,Scratch
SB3,Scratch 3.0 project,Scratch
SBH,Header,ScriptBasic
SBV,Superbase RDBMS form definition data,Superbase (database)
SBX,For experimental extensions to Scratch,Used by scratchX (scratchx.org)
SCALA,Scala source code file,Scala (programming language)
SCM,Scheme source code file,
SCR,Screen Protector file,Windows Screen Protector
SCSS,Sass stylesheet language,
SDF,SQL Server Compact database file,Microsoft SQL Server Compact
SD7,Seed7 source file,Seed7 interpreter and compiler
SDNN,Encrypted Story Distiller Nursery Notes Database,"Story Distiller, Series Distiller"
SDS,Self Defining Structure provides for N-dimensional very large datasets using HHCode,geographic information systems and relational database management systems
SDTS,Spatial Data Transfer Standard,
SEC,Secret key ring file,Pretty Good Privacy RSA System
SED,Self extraction directive file,IExpress
SEQ,Video,Tiertex video sequence
SERIES,Encrypted Series Distiller database,Series Distiller
SF,JAR Digital Signature,
SFB,Configuration file,emacs
SFH,backup file for Strike Force Heroes (SFH) on Steam,
SFX,SFX (self-extracting archives) script,RAR
SH,Unix shell script,Unix shell interpreter
SHAR,Shell self-extracting archive,UNSHAR (Unix)
SHTM,SSI-enabled HTM file,Server Side Includes
SHTML,SSI-enabled HTML file,Server Side Includes
SHX,"Shape entities
ESRI shapefile","AutoCAD
ArcGIS"
SIC,S.I.C.K. Source File,
SIG,Signature file,"gpg, PopMail, ThunderByte AntiVirus"
SL,S-Lang source code file,
SLDASM,SolidWorks assembly,SolidWorks
SLDPRT,SolidWorks part,SolidWorks
SM,SMALLTALK source code file,
SMCLVL,Secret Maryo Chronicles Level,Secret Maryo Chronicles
SMK,Smacker video Format (RAD Video),
SNO,SNOBOL4 source code file,
SO,"shared object, dynamically linked library","Unix, Linux"
SPF,data,SQR Portable Format
SPIFF,Still Picture Interchange File Format,
SPIN,Spin source file,Parallax Propeller Microcontrollers
SPS,SPSS program file (text),"PSPP, SPSS"
SPT,SPITBOL source code file,
SPV,SPIR-V binary file,"Vulkan, Khronos Group"
SPX,Ogg Speex bitstream,Xiph.Org Foundation
SPZ,Crestron SIMPL Windows compiled program (ZIP format),Crestron SIMPL Windows
SQL,Structured Query Language,Any SQL database
SRT,SubRip Subtitle file,Most media players
SRX,Series Distiller update XML file,Series Distiller
SSC,Celestia Solar System Catalog file,Celestia
SSC,Stellarium Script,Stellarium
ST,SMALLTALK source code file,Little Smalltalk
ST,Structured text file,
STC,Celestia Star Catalog file,Celestia
STC,OpenOffice.org XML spreadsheet template,OpenOffice.org Calc
STD,OpenOffice.org XML drawing template,OpenOffice.org Draw
STEP,Standard for the Exchange of Product Data,universal format for CAD exchange per ISO 10303
STI,OpenOffice.org XML presentation template,OpenOffice.org Impress
STK,Stockholm multiple sequence alignment,"Bioinformatics tools eg HMMER, Xrate, Jalview"
STL,surface geometry of a three-dimensional object,software by 3D Systems
STM,SSI-enabled HTML file,Server Side Includes
STO,Stockholm multiple sequence alignment,"Bioinformatics tools eg HMMER, Xrate, Jalview"
STORY,Encrypted Story Distiller database,Story Distiller
STP,Standard for the Exchange of Product Data,universal format for CAD exchange per ISO 10303
STW,OpenOffice.org XML text document template,OpenOffice.org Writer
STX,Story Distiller update XML,Story Distiller
SUR,"Surface topography (in native ""SURF"" format)",MountainsMap
SVC,Represents the ServiceHost instance hosted by Internet Information Services,Windows Communication Foundation
SVELTE,Svelte source code,Svelte
SVG,Scalable Vector Graphics,
SWF,Shockwave Flash,"Macromedia, Adobe Flash Player"
SWG,SWIG source code,SWIG
SWIFT,Swift source code,Swift (programming language)
SWM,split Windows Imaging Format,"ImageX, DISM, 7-Zip, wimlib"
SXC,OpenOffice.org XML spreadsheet,OpenOffice.org Calc
SXD,OpenOffice.org XML drawing,OpenOffice.org Draw
SXG,OpenOffice.org XML master document,
SXI,OpenOffice.org XML presentation,OpenOffice.org Impress
SXM,OpenOffice.org XML formula,OpenOffice.org Calc
SXP,3DS Process file,3D Studio
SXW,OpenOffice.org XML text document,OpenOffice.org Writer
SYLK,Symbolic Link (SYLK) file,Windows
SYMBOLICLINK,Replace file for a symbolic link,Unix-like OSs
TAK,"Audio codec, Lossless audio file format","Winamp (+Plugin), foobar2000 (+Plugin), Media Player Classic – BE"
TAR,tar archive,tar and other file archivers with support
TAZ,tar archive compressed with compress,tar and other file archivers with support
TB2,tar archive compressed with Bzip2,tar and other file archivers with support
TBZ,,
TBZ2,,
TC,Theme Colour file,Saturn CMS
TER,Terragen heightmap file,Terragen scenery generator
TGA,Truevision Advanced Raster Graphics Adapter image,
TGT,Target configuration file,Target active security software
TGZ,tar archive compressed with gzip,tar and other file archivers with support
THM,Thumbnail File,"GoPro, Android, some versions of iOS (Accessible via computer or Jailbreak, on User/Media/PhotoData/Metadata/DCIM)"
TIF,See TIFF,
TIFF,Tag Image File Format image,
TLB,Type library,A binary file with information about a COM or DCOM object so other applications can use it at runtime. Created by Visual C++ or Visual Studio. Used by many Windows applications.
TLZ,tar archive compressed with LZMA,tar and other file archivers with support
TMP,Temporary file,
TORRENT,Torrent file,BitTorrent clients (various)
TQL,The quest lessons,TheQuest
TS,MPEG transport stream,"Video broadcasting, digital video cameras"
TS,TypeScript,
TSCN,Godot Engine Text Scene file,Godot Engine
TSV,Tab-separated values,
TTC,TrueType Font collection,
TTF,TrueType Font file,
TWB,Tableau Software Workbook file,
TXT,Text file,
TXZ,tar archive compressed with xz,tar and other file archivers with support
TZ2,Same as TBZ,
TZST,tar archive compressed with Zstandard,
UI,Espire source code file,Geoworks UI Compiler Geos
UI,Qt Designer's UI File,Trolltech Qt Designer
UMP,Umple UML Programming Language Format,Umple
UNV,"Text file containing finite elements nodal coordinates and more, see notes",Originally used by SDRC for its I-deas software; a lot of simulation software uses it today
UOS,Uniform Office Format spreadsheet,
UOT,Uniform Office Format text,
UPD,Update file for Storage,
UPS,ROM patch file,
URL,Remote file shortcut,Microsoft Windows
USDZ,Augmented Reality (AR) File,"Apple, Pixar"
UST,Vocal synthesis track data,UTAU
USTX,Vocal synthesis track data,OpenUtau
UT!,datafile,uTorrent
UXF,UML Exchange Format,
V,Coq source file,
V,Verilog source file,
V3,Victoria 3 save game file,Victoria 3
V4P,vvvv patch,vvvv
V64,ROM image from an N64 cartridge,"DoctorV64, Doctor V64 junior, Project 64 and other N64 emulators"
VB,Visual Basic .Net source file,Visual Basic .NET
VBOX,virtual machine settings file (in XML format),VirtualBox
VBOX-EXTPACK,VirtualBox extension package,VirtualBox
VBPROJ,Visual Basic .Net project file,Visual Basic .Net Express and Visual Studio 2003-2010 Project
VBR,Visual Basic Custom Control file,Visual Basic
VBS,VBScript script file,VBScript
VBX,Visual Basic eXtension,Visual Basic
VC,VeraCrypt Disk Encrypted file,Open Source VeraCrypt
VC6,Graphite – 2D and 3D drafting,Ashlar-Vellum
VCLS,VocaListener voice scanner file,VocaListener Plug-in (Vocaloid3)
VDA,Targa bitmap graphics,many raster graphics editors
VDI,Virtual Disk Image,VirtualBox
VDW,Visio web drawing,Microsoft Visio
VDX,Visio XML drawing,Microsoft Visio
VFD,Virtual Floppy Disk,Windows Virtual PC (among others)
VI,Virtual Instrument,LabVIEW
VMCZ,Hyper-V Exported Virtual Machine,Microsoft
VMDK,Virtual Disk file,VMware
VMG,Nokia message file format,Text Message Editor (Nokia PC Suite)
VOB,Video Object,"DVD-R, DVD-RW"
VMX,virtual machine configuration file,VMware
VPK,Valve package,Source engine games
VPM,Garmin Voice Processing Module,
VPP,Visual Paradigm Project,Visual Paradigm for UML
VPR,Vocal synthesizer track data,Vocaloid 5
VQM,Hardware description language,Altera
VRB,LateX Beamer file containing verbatim commands,LaTeX Beamer
VRB,Veeam reversed incremental backup archive,Veeam software
VS,Vellum Solids,Ashlar-Vellum
VSD,Visio drawing,Microsoft Visio
VSDX,Visio drawing,Microsoft Visio
VSM,Visual Simulation Model,VisSim
VSQ,Vocal synthesizer track data,Vocaloid 2
VSQx,Vocal synthesizer track data,"Vocaloid 3, Vocaloid 4"
VST,Truevision Vista graphics,many raster graphics editors
VSTO,Microsoft Office add-in file,Microsoft Visual Studio
VSVNBAK,VisualSVN Server repository backup,VisualSVN Server
VTF,Valve Texture Format file,Valve Corporation
VUE,Visual Understanding Environment map,Visual Understanding Environment
VVVVVV,VVVVVV map file,VVVVVV
WAB,Global address book in versions of Microsoft Windows and shared by Microsoft apps such as Outlook and Outlook Express,Outlook and Outlook Express
WAD,"Default package format for Doom that contains sprites, levels, and game data",Doom and Doom II
WAD,"Package containing Wii Channel data, such as Virtual Console games. It is commonly used in homebrew to install custom channels, and can be installed with a WAD Manager",Nintendo Wii
WAV,Sound format (Microsoft Windows RIFF WAVE),Media Player
WEBM,Royalty-free video/audio container,HTML5
WIN,Game code for GameMaker games,GameMaker
WITNESS_CAMPAIGN,Game save file for The Witness,The Witness
WK1,Spreadsheet,Lotus 1-2-3 version 2.x – Lotus Symphony 1.1+
WK3,Spreadsheet,Lotus 1-2-3 version 3.x
WKS,Spreadsheet,"Lotus 1-2-3 version 1A – Lotus Symphony 1.0, Microsoft Works"
WL,Wolfram Language package,
WLMP,"Windows Live Moviemaker Project, contains paths from where the images/audios/videos of the project are located",Windows Live Movie Maker
WLS,Wolfram Script,Wolfram Language
WMA,Windows Media Audio file Advanced Systems Format,
WMDB,Windows Media Player database,Windows Media Player
WMF,Windows MetaFile vector graphics,
WMV,Windows Media Video file Advanced Systems Format,
WOS,WoS bibliographic reference file,ISI/Clarivate Analytics
WPS,Wii U plugin for aroma,Aroma Environnement (Wii U)
WS,"Whitespace programming language, WonderSwan ROM","Whitespace programming language, WonderSwan ROM"
WSC,WonderSwan Color ROM,
WTX,Text document,
WUHB,Wii U Homebrew Bundle,Aroma environnement (Wii U)
X,LEX source code file,
X,XBasic Source code file,Xbasic
X3D,x3d and xdart Formats,
XAR,Xara graphics file,Files created by Xara Photo & Graphic Designer (formerly Xara Xtreme and Xara Studio); early versions used the extension ART
XAR,eXtensible ARchive,"xar, 7-Zip"
XBRL,eXtensible Business Reporting Language instance file,eXtensible Business Reporting Language
XCF,Gimp image file,GNU Image Manipulation Program
XDM,Directory Manipulator for 32-bit Protected Mode,Xenotech Research Labs
XE,Xenon – for Associative 3D Modeling,Ashlar-Vellum
XEX,Xbox 360 Executable File,
XLR,"Microsoft Works spreadsheet or chart file, very similar to Microsoft Excel's XLS",Microsoft Works
XLS,Microsoft Excel Spreadsheet,Microsoft Excel
XLSB,Microsoft Excel 2007 Binary Workbook (BIFF12)(Spreadsheets),Microsoft Excel 2007 (see Microsoft Office 2007 file extensions)
XLSM,Microsoft Excel 2007 Macro-Enabled Workbook (Spreadsheets),Microsoft Excel 2007 (see Microsoft Office 2007 file extensions)
XLSX,Office Open XML Workbook (Spreadsheets),Microsoft Excel 2007 (see Microsoft Office 2007 file extensions)
XM,FastTracker 2 extended module,"AWAVE, Mod4Win, FastTracker, ImpulseTracker"
XML,eXtensible Markup Language file,
XMF,eXtensible Music Format,
XP,eXtended Pattern,FastTracker 2
XPL,X-Plane system file,Laminar Research
XPS,Open XML Paper Specification / OpenXPS,Open standard document format initially created by Microsoft and similar in concept to Adobe PDF files
XSD,XML schema description,
XSF,data,Microsoft InfoPath file
XSL,XSL Stylesheet,
XSLT,XSLT file,
XSN,Microsoft InfoPath template,Microsoft InfoPath
XSPF,XML Sharable Playlist Format,
XX,XX-encoded file (ASCII),XXDECODE
XXE,XX-encoded file (ASCII),XXDECODE
XXX,Singer Embroidery Format,Embroidermodder
XYZ,Molecular coordinates,"XMol, RasMol"
XZ,a lossless data compression file format incorporating the LZMA/LZMA2 compression algorithms.,xz
Y,YACC grammar file,Yacc/Bison
YML,YML markup file,domain specific language output to XML
YAML,YAML source file,YAML (data serialization language)
ZIP,ZIP (file format),PKZip – WinZip – Mac OS X
ZRX,REXX scripting language for ZOC_(software),ZOC terminal emulator
ZS,Script for Minecraft mod MineTweaker and CraftTweaker,Minecraft Mods
ZST,"ZStandard, a lossless data compression file format.",zstd
"@

$CsvData = $csvContent | ConvertFrom-Csv

$Path = $null
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.CheckedExtensions = @{}
# $currentUserIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
# $currentUserIdentityReference = $currentUserIdentity.User # This gives the SecurityIdentifier (SID)
# $currentUserNTAccount = $currentUserIdentity.Name # This gives the NTAccount (Domain\Username)

$Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
$Runspace.ApartmentState = [System.Threading.ApartmentState]::STA
$Runspace.Open()
$RunSpace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
$RunSpace.SessionStateProxy.SetVariable("StatusTextBox", $StatusTextBox)
$RunSpace.SessionStateProxy.SetVariable("ListBoxSelectedExt", $ListBoxSelectedExt)
$PowerShell = [System.Management.Automation.PowerShell]::Create()
$PowerShell.Runspace = $Runspace
    

# [System.Windows.Forms.Control]::Invoke

$Rationals = @(2, 4, 6, 7, 11, 13, 15, 17, 20, 22, 24, 26, 31, 282, 283, 286, 287, 318, 319, 529,
    532, 33423, 33434, 33434, 33437, 33437, 37122, 37122, 37378, 37378, 37381, 37381,
    37382, 37386, 37386, 37387, 37390, 37391, 37397, 37889, 37890, 37892, 41483, 41486,
    41487, 41493, 41988, 42034, 42240, 50714, 50718, 50727, 50729, 50731, 50732, 50734,
    50736, 50737, 50738, 50780, 50935, 51125, 51178, 51179)

$Longs = @(254, 256, 257, 273, 278, 279, 292, 293, 322, 323, 325, 330, 512, 513, 514, 519, 520, 521, 4097,
    4098, 33723, 34665, 34853, 34865, 34866, 34867, 34868, 34869, 37393, 40962, 40963, 40965, 45059,
    45060, 45313, 45569, 45570, 45571, 45572, 45573, 45574, 45575, 45576, 45577, 45578, 45579, 45580,
    45581, 50717, 50719, 50720, 50733, 50829, 50830, 50933, 50937, 50941, 50970, 50974, 50975, 50981,
    51089, 51090, 51091, 51107, 51108, 51110, 52536, 52547, 52554, 52555)

function Get-Folder
{
    $Path = $null

    $CheckedItemsList = @()
    # Iterate through each control on the form
    foreach ($Control in $Form.Controls)
    {
        # Check if the control is a ListView
        if ($Control -is [System.Windows.Forms.ListView])
        {
            # Loop through the checked items in the current ListView
            foreach ($CheckedItem in $Control.CheckedItems)
            {
                $CheckedItemsList += $CheckedItem.Text
            }
        }
    }

    if ($CheckedItemsList -ne "")
    {
        $Include = $CheckedItemsList | ForEach-Object { "*.$_*" }
    }
    else
    {
        $Include = $null
    }

    $OpenFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $OpenFolderDialog.SelectedPath = "D:\Developement\MetaDataFiles"
    $OpenFolderDialog.ShowNewFolderButton = $false
    $OpenFolderDialog.Description = "Select a directory"
    $CancelCheck = $OpenFolderDialog.ShowDialog()

    if ($CancelCheck -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $Path = $OpenFolderDialog.SelectedPath
        return $Path, $Include
    }
    elseif([System.Windows.Forms.DialogResult]::Cancel)
    {
        return $null, $null
    }
}

function Get-Files
{
    $Path = $null

    $TopmostForm = New-Object System.Windows.Forms.Form -Property @{TopMost = $true }

    # Build the filter string for the OpenFileDialog
    $FilterString = ""
    foreach ($Extension in $CsvData)
    {
        if ($Extension.Description.Length -gt 30)
        {
            # Description for the filter
            $FullDescription = $Extension.Description
            $Description = $FullDescription.Substring(0, 30)
        }
        else
        {
            $Description = $Extension.Description
        }
        
        # Extensions (e.g., txt, jpg;png;gif)
        $Exts = $Extension.Ext
        # Formats the extensions for the filter string (e.g., *.txt, *.jpg;*.png;*.gif)
        $FilterPatterns = $Exts | ForEach-Object { "*.$_" } | Out-String -Stream -Width 10 | ForEach-Object { $_.Trim() }

        # Append to the filter string
        $FilterString += "($FilterPatterns) $Description |$FilterPatterns|"
    }

    # Remove the trailing "|"
    $FilterString = $FilterString.TrimEnd('|')

    # Create an instance of the OpenFileDialog
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Title = "Select a File"
    $FileBrowser.InitialDirectory = [Environment]::GetFolderPath("MyComputer")
    # $FileBrowser.InitialDirectory = "D:\Developement\MetaDataFiles"
    $FileBrowser.Filter = $FilterString
    $FileBrowser.FilterIndex = 1 # Set the default filter
    $FileBrowser.Multiselect = $true # Multiple files can be chosen


    $CancelCheck = $FileBrowser.ShowDialog($TopmostForm)

    if ($CancelCheck -eq [System.Windows.Forms.DialogResult]::OK)
    {
        if($FileBrowser.FileNames -like "*\*")
        {
            $Path = $FileBrowser.FileNames
            return $Path, $Include
        }
    }
    elseif([System.Windows.Forms.DialogResult]::Cancel)
    {
        return $null, $null
       
    }
}

$PowerShell.AddScript({
        function Test-HasExifData
        {
            param (
                [Parameter(Mandatory = $true)]
                [string]$ImagePath
            )

            if (@(".bmp", ".gif", ".jpeg", ".jpg", ".exif", ".png", ".tiff", ".tif") -notcontains [System.IO.Path]::GetExtension($ImagePath))
            {
                return $false
            }

            try
            {
                $bitmap = [System.Drawing.Bitmap]::new($ImagePath)
                if ($bitmap.PropertyItems.Count -eq 0)
                {
                    return $true, 0
                } 
                if ($bitmap.PropertyItems.Count -gt 0)
                {
                    return $true, 1
                }
                else
                {
                    return $false
                }
            }
            catch
            {
                # Catch errors if the file is not a valid image or doesn't have accessible properties
                Write-Warning "Could not process '$ImagePath': $($_.Exception.Message)"
                return $false
            }
            finally
            {
                # Dispose of the bitmap object to release file handle
                if ($bitmap)
                {
                    $bitmap.Dispose() | Out-Null
                }
            }
        } }).Invoke()

$PowerShell.AddScript({
        function Get-AllData
        {
            param (
                $Path,
                $Include,
                $StatusTextBox,
                $SyncHash,
                $ListBoxSelectedExt,
                $FindOnly,
                $ShowInaccessable
            )

            $StatusTextBox.Clear()
            $ExifResults = @()
            $Cancel = $false
            $Shell = New-Object -ComObject Shell.Application

            if(($null -ne $Path) -and ($null -ne $Include))
            {
                $StatusTextBox.AppendText("Getting Files/Folders...`r`n")
                $AllFiles = Get-ChildItem -Path $Path -Include $Include -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable InaccessibleItems
            }
            elseif($null -ne $Path)
            {
                $StatusTextBox.AppendText("Getting Files/Folders...`r`n")
                $AllFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable InaccessibleItems
            }
            else
            {
                $Cancel = $true
            }

            if ($Cancel -ne $true)
            {
                if ($FindOnly.Checked -eq $true)
                {
                    # Return the simpler data structure for the "Find Only" case
                    return @{FindOnlyResults = ($AllFiles | Select-Object Name, DirectoryName, Extension); InaccessibleItems = $InaccessibleItems}
                }
                elseif($FindOnly.Checked -ne $true)
                {
                    $ExifResults = foreach($A in $AllFiles)
                    {
                        $StatusTextBox.AppendText("Processing: $($A.Name)`r`n")
                        $StatusTextBox.ScrollToCaret() 
                        $Bitmap = $null
                        $HasExifData = Test-HasExifData -ImagePath $A.FullName

                        if ($HasExifData -eq $false)
                        {
                            [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = "Not Exist"; Group = "Not Exist"; Name = "Not Exist"; Type = "Not Exist"; Value = "Not Exist" }
                        }
                        elseif (($HasExifData[0] -eq $true) -and ($HasExifData[1] -eq 0))
                        {
                            [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = "Viable Exif"; Group = "But"; Name = "No"; Type = "Data"; Value = "Found" }
                        }
                        elseif (($HasExifData[0] -eq $true) -and ($HasExifData[1] -eq 1))
                        {
                            $Bitmap = [System.Drawing.Bitmap]::new($A.FullName)
                            for ($i = 0; $i -lt 53000 ; $i++)
                            {
                                try
                                {
                                    $GetBitmap = $null
                                    $GetBitmap = $Bitmap.GetPropertyItem($i)
                                    if ($null -ne $GetBitmap.Value)
                                    {

                                        if ($Rationals -contains $GetBitmap.ID)
                                        {
                                            $Numerator = [System.BitConverter]::ToUInt32($GetBitmap.Value, 0)
                                            $Denominator = [System.BitConverter]::ToUInt32($GetBitmap.Value, 4)
                                            if ($Denominator -ne 0)
                                            {
                                                $DecimalValue = $Numerator / $Denominator
                                            }
                                            else
                                            {
                                                $DecimalValue = "Unknown Rational"
                                            }
                                        }

                                        if ($Longs -contains $GetBitmap.ID)
                                        {
                                            $SexagesimalValue = $GetBitmap.Value.ToString
                                            $Parts = $SexagesimalValue.Split(' ')
                                            $Degrees = [double]$Parts[0]
                                            $Minutes = [double]$Parts[1]
                                            $Seconds = [double]$Parts[2]
                                            $Direction = $Parts[3]

                                            $DecimalValue = $Degrees + ($Minutes / 60) + ($Seconds / 3600)

                                            if ($Direction -eq 'S' -or $Direction -eq 'W')
                                            {
                                                $DecimalValue = - $DecimalValue
                                            }
                                        }
                                
                                        $Id = [int]::Parse($GetBitmap.Id)
                                        $ExifResults += Get-SwitchResult -GetBitmapId $Id
                                    }
                                }
                                catch
                                {
                                    continue
                                }
                            }    
                        }

                        $Folder = $Shell.NameSpace((Split-Path $A.FullName))
                        $File = $Folder.ParseName((Split-Path $A.FullName -Leaf))

                        for ($i = 0; $i -lt 3000; $i++)
                        {
                            $PropertyName = $Folder.GetDetailsOf($Folder.Items, $i)
                            if ($PropertyName)
                            {
                                $PropertyValue = $Folder.GetDetailsOf($File, $i)
                                if ($PropertyValue)
                                {
                                    [PSCustomObject]@{ Source = "Meta"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Not Exist"; Name = $PropertyName; Type = $PropertyValue.GetType().FullName; Value = $PropertyValue }
                                }
                    
                            }
                        }
                    }

                    $StatusTextBox.AppendText("Complete`r`n")
                    $StatusTextBox.ScrollToCaret() 
                    
                    # The results are now returned to be displayed in a custom form.
                    return @{ExifResults = $ExifResults; InaccessibleItems = $InaccessibleItems}
                }
        
            }
        }
    }).Invoke() # Invoke to add the function to the session state

function Show-ResultsForm {
    param(
        [System.Collections.IDictionary]$ResultsData,
        [System.Windows.Forms.CheckBox]$ShowInaccessableCheckbox
    )

    $resultsForm = New-Object System.Windows.Forms.Form -Property @{
        Text          = "Scan Results"
        Size          = New-Object System.Drawing.Size(1200, 800)
        StartPosition = "CenterScreen"
    }

    $dataGridView = New-Object System.Windows.Forms.DataGridView -Property @{
        Dock                      = [System.Windows.Forms.DockStyle]::Fill
        AllowUserToAddRows        = $false
        ReadOnly                  = $true
        AutoSizeColumnsMode       = "Fill"
        AutoGenerateColumns       = $false # We will define columns manually for reliability
    }

    # Set cell style properties AFTER creating the object for reliability
    $dataGridView.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    # This ensures rows will resize to fit the content of wrapped cells.
    $dataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells

    # --- Manually Define Columns ---
    $colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Source"; HeaderText = "Source" }
    $colFilename = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Filename"; HeaderText = "Filename" }
    $colPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Path"; HeaderText = "Path" }
    $colID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "ID"; HeaderText = "ID" }
    $colGroup = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Group"; HeaderText = "Group" }
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Name"; HeaderText = "Name" }
    $colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Type"; HeaderText = "Type" }
    $colValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Value"; HeaderText = "Value" }

    $dataGridView.Columns.AddRange($colSource, $colFilename, $colPath, $colID, $colGroup, $colName, $colType, $colValue)

    # --- Event Handler for ToolTips ---
    # This will show the full cell content when the user hovers over a cell.
    $dataGridView.add_CellToolTipTextNeeded({
        param($sender, $e)
        # Ensure we are not on the header row
        if ($e.RowIndex -gt -1 -and $e.ColumnIndex -gt -1) {
            $e.ToolTipText = $sender.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
        }
    })

    # --- Manually Add Data Rows ---
    if ($ResultsData.ExifResults) {
        # Suspend layout for performance during bulk add
        $dataGridView.SuspendLayout()
        foreach ($item in $ResultsData.ExifResults) {
            $dataGridView.Rows.Add($item.Source, $item.Filename, $item.Path, $item.ID, $item.Group, $item.Name, $item.Type, $item.Value) | Out-Null
        }
        # Resume layout after adding all rows
        $dataGridView.ResumeLayout()
    }

    # --- Bottom Panel for Controls ---
    $bottomPanel = New-Object System.Windows.Forms.Panel -Property @{
        Dock   = [System.Windows.Forms.DockStyle]::Bottom
        Height = 50
    }

    $buttonSelectFile = New-Object System.Windows.Forms.Button -Property @{
        Text     = "Select Export File"
        Location = New-Object System.Drawing.Point(10, 10)
        Size     = New-Object System.Drawing.Size(130, 30)
    }

    $txtExportPath = New-Object System.Windows.Forms.TextBox -Property @{
        Text     = (Join-Path $PSScriptRoot "AllFilesData.csv")
        Location = New-Object System.Drawing.Point(150, 12)
        Size     = New-Object System.Drawing.Size(400, 25)
        ReadOnly = $true
    }

    $buttonExport = New-Object System.Windows.Forms.Button -Property @{
        Text     = "Export to CSV"
        Location = New-Object System.Drawing.Point(560, 10)
        Size     = New-Object System.Drawing.Size(120, 30)
    }

    $buttonClose = New-Object System.Windows.Forms.Button -Property @{
        Text         = "Close"
        Anchor       = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
        Location     = New-Object System.Drawing.Point(1090, 10)
        Size         = New-Object System.Drawing.Size(90, 30)
        DialogResult = [System.Windows.Forms.DialogResult]::OK
    }

    # --- Event Handlers for Results Form ---
    $buttonSelectFile.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{
            Filter      = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            Title       = "Select a file to save the results"
            FileName    = "AllFilesData.csv"
        }
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtExportPath.Text = $saveFileDialog.FileName
        }
        $saveFileDialog.Dispose()
    })

    $buttonExport.Add_Click({
        try {
            $ResultsData.ExifResults | Export-Csv -Path $txtExportPath.Text -NoTypeInformation -Force
            [System.Windows.Forms.MessageBox]::Show("Successfully exported results to:`n$($txtExportPath.Text)", "Export Complete", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred during export:`n$($_.Exception.Message)", "Export Error", "OK", "Error")
        }
    })

    $bottomPanel.Controls.AddRange(@($buttonSelectFile, $txtExportPath, $buttonExport, $buttonClose))
    $resultsForm.Controls.AddRange(@($dataGridView, $bottomPanel))

    # Show inaccessible items in a separate grid view if they exist
    if ($ShowInaccessableCheckbox.Checked -and $ResultsData.InaccessibleItems.Count -gt 0) {
        $ResultsData.InaccessibleItems | Select-Object TargetObject, CategoryInfo | Out-GridView -Title "Inaccessible Folders" -Wait
    }

    $resultsForm.ShowDialog($Form) | Out-Null
    $resultsForm.Dispose()
}

function Show-FindOnlyResultsForm {
    param(
        [System.Collections.IDictionary]$ResultsData,
        [System.Windows.Forms.CheckBox]$ShowInaccessableCheckbox
    )

    $resultsForm = New-Object System.Windows.Forms.Form -Property @{
        Text          = "Find Only Results"
        Size          = New-Object System.Drawing.Size(1000, 700)
        StartPosition = "CenterScreen"
    }

    $dataGridView = New-Object System.Windows.Forms.DataGridView -Property @{
        Dock                      = [System.Windows.Forms.DockStyle]::Fill
        AllowUserToAddRows        = $false
        ReadOnly                  = $true
        AutoSizeColumnsMode       = "Fill"
        AutoGenerateColumns       = $false
    }

    # --- Manually Define Columns for "Find Only" ---
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Name"; HeaderText = "Name" }
    $colDirectoryName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "DirectoryName"; HeaderText = "Directory" }
    $colExtension = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Extension"; HeaderText = "Extension" }
    $dataGridView.Columns.AddRange($colName, $colDirectoryName, $colExtension)

    # --- Manually Add Data Rows ---
    if ($ResultsData.FindOnlyResults) {
        $dataGridView.SuspendLayout()
        foreach ($item in $ResultsData.FindOnlyResults) {
            $dataGridView.Rows.Add($item.Name, $item.DirectoryName, $item.Extension) | Out-Null
        }
        $dataGridView.ResumeLayout()
    }

    # --- Bottom Panel for Controls ---
    $bottomPanel = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Bottom; Height = 50 }
    $buttonSelectFile = New-Object System.Windows.Forms.Button -Property @{ Text = "Select Export File"; Location = New-Object System.Drawing.Point(10, 10); Size = New-Object System.Drawing.Size(130, 30) }
    $txtExportPath = New-Object System.Windows.Forms.TextBox -Property @{ Text = (Join-Path $PSScriptRoot "FindOnlyResults.csv"); Location = New-Object System.Drawing.Point(150, 12); Size = New-Object System.Drawing.Size(400, 25); ReadOnly = $true }
    $buttonExport = New-Object System.Windows.Forms.Button -Property @{ Text = "Export to CSV"; Location = New-Object System.Drawing.Point(560, 10); Size = New-Object System.Drawing.Size(120, 30) }
    $buttonClose = New-Object System.Windows.Forms.Button -Property @{ Text = "Close"; Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right); Location = New-Object System.Drawing.Point(900, 10); Size = New-Object System.Drawing.Size(90, 30); DialogResult = [System.Windows.Forms.DialogResult]::OK }

    # --- Event Handlers for FindOnly Form ---
    $buttonSelectFile.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{ Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"; Title = "Select a file to save the results"; FileName = "FindOnlyResults.csv" }
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtExportPath.Text = $saveFileDialog.FileName }
        $saveFileDialog.Dispose()
    })

    $buttonExport.Add_Click({
        try {
            $ResultsData.FindOnlyResults | Export-Csv -Path $txtExportPath.Text -NoTypeInformation -Force
            [System.Windows.Forms.MessageBox]::Show("Successfully exported results to:`n$($txtExportPath.Text)", "Export Complete", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred during export:`n$($_.Exception.Message)", "Export Error", "OK", "Error")
        }
    })

    $bottomPanel.Controls.AddRange(@($buttonSelectFile, $txtExportPath, $buttonExport, $buttonClose))
    $resultsForm.Controls.AddRange(@($dataGridView, $bottomPanel))

    # Show inaccessible items in a separate grid view if they exist
    if ($ShowInaccessableCheckbox.Checked -and $ResultsData.InaccessibleItems.Count -gt 0) {
        $ResultsData.InaccessibleItems | Select-Object TargetObject, CategoryInfo | Out-GridView -Title "Inaccessible Folders" -Wait
    }

    $resultsForm.ShowDialog($Form) | Out-Null
    $resultsForm.Dispose()
}

function Get-FilteredExtensions
{
    param (
        [string]$CurrentLetter,
        [object[]]$CsvData
    )

    if($CurrentLetter -match "09")
    {
        $CsvLetter = "0|^1|^2|^3|^4|^5|^6|^7|^8|^9"
        $SyncHash.FilteredData = $CsvData | Where-Object { $_.Ext -match "^$CsvLetter" }
    }
    else
    {
        $SyncHash.FilteredData = $CsvData | Where-Object { $_.Ext -match "^$CurrentLetter" }
    }

    $MatchingListView = $Form.Controls | Where-Object { $_.Name -match "ListView$($CurrentLetter)" }

    foreach ($row in $SyncHash.FilteredData)
    {
        $item = New-Object System.Windows.Forms.ListViewItem($row.Ext) # Use your first column header
        $item.SubItems.Add($row.Description) # Use your second column header
        $item.SubItems.Add($row.'Used by') # Use your third column header
        # Add more subitems for additional columns as needed

        $MatchingListView.Items.Add($item) | Out-Null # Use Out-Null to suppress output
    }
}

$PowerShell.AddScript({
        function Get-SwitchResult
        {
            param ([string]$GetBitmapId)
            switch ($GetBitmapId)
            {
                #1 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Iop"; Name = "InteroperabilityIndex"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$"}}
                1 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSLatitudeRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                3 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSLongitudeRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                8 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSSatellites"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                9 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSStatus"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                10 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSMeasureMode"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                11 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProcessingSoftware"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                12 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSSpeedRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                14 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSTrackRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                16 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSImgDirectionRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                18 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSMapDatum"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                19 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDestLatitudeRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                21 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDestLongitudeRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                23 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDestBearingRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                25 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDestDistanceRef"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                29 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDateStamp"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                269 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DocumentName"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                270 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageDescription"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                271 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Make"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                272 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Model"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                285 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PageName"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                305 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Software"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                306 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DateTime"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                315 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Artist"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                316 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "HostComputer"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                333 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "InkNames"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                337 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TargetPrinter"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                4096 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Iop"; Name = "RelatedImageFileFormat"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                32781 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageID"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                33432 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Copyright"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                34852 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SpectralSensitivity"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                34852 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SpectralSensitivity"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                36867 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DateTimeOriginal"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                36867 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "DateTimeOriginal"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                36868 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "DateTimeDigitized"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                36880 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "OffsetTime"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                36881 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "OffsetTimeOriginal"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                36882 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "OffsetTimeDigitized"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                37394 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SecurityClassification"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                37395 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageHistory"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                37520 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SubSecTime"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                37521 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SubSecTimeOriginal"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                37522 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SubSecTimeDigitized"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                40964 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "RelatedSoundFile"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42016 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ImageUniqueID"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42032 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "CameraOwnerName"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42033 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "BodySerialNumber"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42035 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "LensMake"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42036 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "LensModel"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42037 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "LensSerialNumber"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42038 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ImageTitle"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42039 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Photographer"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42040 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ImageEditor"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42041 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "CameraFirmware"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42042 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "RAWDevelopingSoftware"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42043 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ImageEditingSoftware"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                42044 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "MetadataEditingSoftware"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                45056 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFVersion"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                45058 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFImageList"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                50708 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "UniqueCameraModel"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                50735 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CameraSerialNumber"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                50971 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PreviewDateTime"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                51081 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ReelName"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                51105 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CameraLabel"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                51182 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "EnhanceParams"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                52526 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SemanticName"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                52528 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SemanticInstanceID"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                52552 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileGroupName"; Type = "Ascii"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                0 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSVersionID"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                5 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSAltitudeRef"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                336 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DotRange"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                343 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ClipPath"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                700 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XMLPacket"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                33422 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CFAPattern"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                34377 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageResources"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                37398 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TIFFEPStandardID"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                40091 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XPTitle"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                40092 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XPComment"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                40093 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XPAuthor"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                40094 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XPKeywords"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                40095 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XPSubject"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50706 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DNGVersion"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50707 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DNGBackwardVersion"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50709 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "LocalizedCameraModel"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50710 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CFAPlaneColor"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50740 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DNGPrivateData"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50781 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "RawDataUniqueID"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50827 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OriginalRawFileName"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50931 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CameraCalibrationSignature"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50932 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileCalibrationSignature"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50934 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "AsShotProfileName"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50936 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileName"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50942 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileCopyright"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50966 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PreviewApplicationName"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50967 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PreviewApplicationVersion"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50968 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PreviewSettingsName"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                50969 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PreviewSettingsDigest"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                51043 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TimeCodes"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                51111 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "NewRawImageDigest"; Type = "Byte"; Value = $GetBitmap.Value[0] } }
                27 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSProcessingMethod"; Type = "Comment"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                28 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSAreaInformation"; Type = "Comment"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                37510 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "UserComment"; Type = "Comment"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$" } }
                51041 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "NoiseProfile"; Type = "Double"; Value = [System.BitConverter]::ToSingle($GetBitmap.Value, 0) } }
                51112 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "RawToPreviewGain"; Type = "Double"; Value = [System.BitConverter]::ToSingle($GetBitmap.Value, 0) } }
                50938 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileHueSatMapData1"; Type = "Float"; Value = [System.BitConverter]::ToSingle($GetBitmap.Value, 0) } }
                50939 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileHueSatMapData2"; Type = "Float"; Value = [System.BitConverter]::ToSingle($GetBitmap.Value, 0) } }
                50940 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileToneCurve"; Type = "Float"; Value = [System.BitConverter]::ToSingle($GetBitmap.Value, 0) } }
                50982 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileLookTableData"; Type = "Float"; Value = [System.BitConverter]::ToSingle($GetBitmap.Value, 0) } }
                52537 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileHueSatMapData3"; Type = "Float"; Value = [System.BitConverter]::ToSingle($GetBitmap.Value, 0) } }
                52553 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JXLDistance"; Type = "Float"; Value = [System.BitConverter]::ToSingle($GetBitmap.Value, 0) } }
                254 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "NewSubfileType"; Type = "Long"; Value = $DecimalValue } }
                256 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageWidth"; Type = "Long"; Value = $DecimalValue } }
                257 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageLength"; Type = "Long"; Value = $DecimalValue } }
                273 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "StripOffsets"; Type = "Long"; Value = $DecimalValue } }
                278 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "RowsPerStrip"; Type = "Long"; Value = $DecimalValue } }
                279 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "StripByteCounts"; Type = "Long"; Value = $DecimalValue } }
                292 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "T4Options"; Type = "Long"; Value = $DecimalValue } }
                293 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "T6Options"; Type = "Long"; Value = $DecimalValue } }
                322 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TileWidth"; Type = "Long"; Value = $DecimalValue } }
                323 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TileLength"; Type = "Long"; Value = $DecimalValue } }
                325 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TileByteCounts"; Type = "Long"; Value = $DecimalValue } }
                330 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SubIFDs"; Type = "Long"; Value = $DecimalValue } }
                512 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGProc"; Type = "Long"; Value = $DecimalValue } }
                513 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGInterchangeFormat"; Type = "Long"; Value = $DecimalValue } }
                514 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGInterchangeFormatLength"; Type = "Long"; Value = $DecimalValue } }
                519 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGQTables"; Type = "Long"; Value = $DecimalValue } }
                520 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGDCTables"; Type = "Long"; Value = $DecimalValue } }
                521 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGACTables"; Type = "Long"; Value = $DecimalValue } }
                4097 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Iop"; Name = "RelatedImageWidth"; Type = "Long"; Value = $DecimalValue } }
                4098 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Iop"; Name = "RelatedImageLength"; Type = "Long"; Value = $DecimalValue } }
                33723 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "IPTCNAA"; Type = "Long"; Value = $DecimalValue } }
                34665 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ExifTag"; Type = "Long"; Value = $DecimalValue } }
                34853 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "GPSTag"; Type = "Long"; Value = $DecimalValue } }
                34865 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "StandardOutputSensitivity"; Type = "Long"; Value = $DecimalValue } }
                34866 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "RecommendedExposureIndex"; Type = "Long"; Value = $DecimalValue } }
                34867 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ISOSpeed"; Type = "Long"; Value = $DecimalValue } }
                34868 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ISOSpeedLatitudeyyy"; Type = "Long"; Value = $DecimalValue } }
                34869 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ISOSpeedLatitudezzz"; Type = "Long"; Value = $DecimalValue } }
                37393 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageNumber"; Type = "Long"; Value = $DecimalValue } }
                40962 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "PixelXDimension"; Type = "Long"; Value = $DecimalValue } }
                40963 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "PixelYDimension"; Type = "Long"; Value = $DecimalValue } }
                40965 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "InteroperabilityTag"; Type = "Long"; Value = $DecimalValue } }
                45059 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFImageUIDList"; Type = "Long"; Value = $DecimalValue } }
                45060 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFTotalFrames"; Type = "Long"; Value = $DecimalValue } }
                45313 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFIndividualNum"; Type = "Long"; Value = $DecimalValue } }
                45569 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFPanOrientation"; Type = "Long"; Value = $DecimalValue } }
                45570 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFPanOverlapH"; Type = "Long"; Value = $DecimalValue } }
                45571 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFPanOverlapV"; Type = "Long"; Value = $DecimalValue } }
                45572 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFBaseViewpointNum"; Type = "Long"; Value = $DecimalValue } }
                45573 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFConvergenceAngle"; Type = "Long"; Value = $DecimalValue } }
                45574 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFBaselineLength"; Type = "Long"; Value = $DecimalValue } }
                45575 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFVerticalDivergence"; Type = "Long"; Value = $DecimalValue } }
                45576 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFAxisDistanceX"; Type = "Long"; Value = $DecimalValue } }
                45577 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFAxisDistanceY"; Type = "Long"; Value = $DecimalValue } }
                45578 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFAxisDistanceZ"; Type = "Long"; Value = $DecimalValue } }
                45579 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFYawAngle"; Type = "Long"; Value = $DecimalValue } }
                45580 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFPitchAngle"; Type = "Long"; Value = $DecimalValue } }
                45581 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFRollAngle"; Type = "Long"; Value = $DecimalValue } }
                50717 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "WhiteLevel"; Type = "Long"; Value = $DecimalValue } }
                50719 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DefaultCropOrigin"; Type = "Long"; Value = $DecimalValue } }
                50720 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DefaultCropSize"; Type = "Long"; Value = $DecimalValue } }
                50733 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BayerGreenSplit"; Type = "Long"; Value = $DecimalValue } }
                50829 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ActiveArea"; Type = "Long"; Value = $DecimalValue } }
                50830 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "MaskedAreas"; Type = "Long"; Value = $DecimalValue } }
                50933 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ExtraCameraProfiles"; Type = "Long"; Value = $DecimalValue } }
                50937 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileHueSatMapDims"; Type = "Long"; Value = $DecimalValue } }
                50941 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileEmbedPolicy"; Type = "Long"; Value = $DecimalValue } }
                50970 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PreviewColorSpace"; Type = "Long"; Value = $DecimalValue } }
                50974 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SubTileBlockSize"; Type = "Long"; Value = $DecimalValue } }
                50975 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "RowInterleaveFactor"; Type = "Long"; Value = $DecimalValue } }
                50981 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileLookTableDims"; Type = "Long"; Value = $DecimalValue } }
                51089 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OriginalDefaultFinalSize"; Type = "Long"; Value = $DecimalValue } }
                51090 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OriginalBestQualityFinalSize"; Type = "Long"; Value = $DecimalValue } }
                51091 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OriginalDefaultCropSize"; Type = "Long"; Value = $DecimalValue } }
                51107 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileHueSatMapEncoding"; Type = "Long"; Value = $DecimalValue } }
                51108 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileLookTableEncoding"; Type = "Long"; Value = $DecimalValue } }
                51110 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DefaultBlackRender"; Type = "Long"; Value = $DecimalValue } }
                52536 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "MaskSubArea"; Type = "Long"; Value = $DecimalValue } }
                52547 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ColumnInterleaveFactor"; Type = "Long"; Value = $DecimalValue } }
                52554 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JXLEffort"; Type = "Long"; Value = $DecimalValue } }
                52555 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JXLDecodeSpeed"; Type = "Long"; Value = $DecimalValue } }
                2 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSLatitude"; Type = "Rational"; Value = $DecimalValue } }
                4 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSLongitude"; Type = "Rational"; Value = $DecimalValue } }
                6 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSAltitude"; Type = "Rational"; Value = $DecimalValue } }
                7 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSTimeStamp"; Type = "Rational"; Value = $DecimalValue } }
                #11 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDOP"; Type = "Rational"; Value = $DecimalValue}}
                13 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSSpeed"; Type = "Rational"; Value = $DecimalValue } }
                15 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSTrack"; Type = "Rational"; Value = $DecimalValue } }
                17 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSImgDirection"; Type = "Rational"; Value = $DecimalValue } }
                20 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDestLatitude"; Type = "Rational"; Value = $DecimalValue } }
                22 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDestLongitude"; Type = "Rational"; Value = $DecimalValue } }
                24 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDestBearing"; Type = "Rational"; Value = $DecimalValue } }
                26 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDestDistance"; Type = "Rational"; Value = $DecimalValue } }
                31 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSHPositioningError"; Type = "Rational"; Value = $DecimalValue } }
                282 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XResolution"; Type = "Rational"; Value = $DecimalValue } }
                283 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "YResolution"; Type = "Rational"; Value = $DecimalValue } }
                286 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XPosition"; Type = "Rational"; Value = $DecimalValue } }
                287 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "YPosition"; Type = "Rational"; Value = $DecimalValue } }
                318 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "WhitePoint"; Type = "Rational"; Value = $DecimalValue } }
                319 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PrimaryChromaticities"; Type = "Rational"; Value = $DecimalValue } }
                529 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "YCbCrCoefficients"; Type = "Rational"; Value = $DecimalValue } }
                532 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ReferenceBlackWhite"; Type = "Rational"; Value = $DecimalValue } }
                33423 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BatteryLevel"; Type = "Rational"; Value = $DecimalValue } }
                33434 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ExposureTime"; Type = "Rational"; Value = $DecimalValue } }
                33434 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ExposureTime"; Type = "Rational"; Value = $DecimalValue } }
                33437 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "FNumber"; Type = "Rational"; Value = $DecimalValue } }
                33437 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FNumber"; Type = "Rational"; Value = $DecimalValue } }
                37122 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CompressedBitsPerPixel"; Type = "Rational"; Value = $DecimalValue } }
                37122 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "CompressedBitsPerPixel"; Type = "Rational"; Value = $DecimalValue } }
                37378 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ApertureValue"; Type = "Rational"; Value = $DecimalValue } }
                37378 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ApertureValue"; Type = "Rational"; Value = $DecimalValue } }
                37381 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "MaxApertureValue"; Type = "Rational"; Value = $DecimalValue } }
                37381 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "MaxApertureValue"; Type = "Rational"; Value = $DecimalValue } }
                37382 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SubjectDistance"; Type = "Rational"; Value = $DecimalValue } }
                37386 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "FocalLength"; Type = "Rational"; Value = $DecimalValue } }
                37386 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FocalLength"; Type = "Rational"; Value = $DecimalValue } }
                37387 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "FlashEnergy"; Type = "Rational"; Value = $DecimalValue } }
                37390 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "FocalPlaneXResolution"; Type = "Rational"; Value = $DecimalValue } }
                37391 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "FocalPlaneYResolution"; Type = "Rational"; Value = $DecimalValue } }
                37397 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ExposureIndex"; Type = "Rational"; Value = $DecimalValue } }
                37889 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Humidity"; Type = "Rational"; Value = $DecimalValue } }
                37890 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Pressure"; Type = "Rational"; Value = $DecimalValue } }
                37892 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Acceleration"; Type = "Rational"; Value = $DecimalValue } }
                41483 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FlashEnergy"; Type = "Rational"; Value = $DecimalValue } }
                41486 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FocalPlaneXResolution"; Type = "Rational"; Value = $DecimalValue } }
                41487 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FocalPlaneYResolution"; Type = "Rational"; Value = $DecimalValue } }
                41493 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ExposureIndex"; Type = "Rational"; Value = $DecimalValue } }
                41988 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "DigitalZoomRatio"; Type = "Rational"; Value = $DecimalValue } }
                42034 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "LensSpecification"; Type = "Rational"; Value = $DecimalValue } }
                42240 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Gamma"; Type = "Rational"; Value = $DecimalValue } }
                50714 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BlackLevel"; Type = "Rational"; Value = $DecimalValue } }
                50718 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DefaultScale"; Type = "Rational"; Value = $DecimalValue } }
                50727 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "AnalogBalance"; Type = "Rational"; Value = $DecimalValue } }
                50729 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "AsShotWhiteXY"; Type = "Rational"; Value = $DecimalValue } }
                50731 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BaselineNoise"; Type = "Rational"; Value = $DecimalValue } }
                50732 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BaselineSharpness"; Type = "Rational"; Value = $DecimalValue } }
                50734 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "LinearResponseLimit"; Type = "Rational"; Value = $DecimalValue } }
                50736 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "LensInfo"; Type = "Rational"; Value = $DecimalValue } }
                50737 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ChromaBlurRadius"; Type = "Rational"; Value = $DecimalValue } }
                50738 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "AntiAliasStrength"; Type = "Rational"; Value = $DecimalValue } }
                50780 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BestQualityScale"; Type = "Rational"; Value = $DecimalValue } }
                50935 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "NoiseReductionApplied"; Type = "Rational"; Value = $DecimalValue } }
                51125 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DefaultUserCrop"; Type = "Rational"; Value = $DecimalValue } }
                51178 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DepthNear"; Type = "Rational"; Value = $DecimalValue } }
                51179 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DepthFar"; Type = "Rational"; Value = $DecimalValue } }
                30 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "GPSInfo"; Name = "GPSDifferential"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                255 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SubfileType"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                258 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BitsPerSample"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                259 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Compression"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                262 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PhotometricInterpretation"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                263 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Thresholding"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                264 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CellWidth"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                265 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CellLength"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                266 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "FillOrder"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                274 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Orientation"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                277 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SamplesPerPixel"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                284 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PlanarConfiguration"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                290 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "GrayResponseUnit"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                291 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "GrayResponseCurve"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                296 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ResolutionUnit"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                297 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PageNumber"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                301 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TransferFunction"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                317 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Predictor"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                320 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ColorMap"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                321 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "HalftoneHints"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                324 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TileOffsets"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                332 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "InkSet"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                334 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "NumberOfInks"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                338 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ExtraSamples"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                339 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SampleFormat"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                340 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SMinSampleValue"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                341 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SMaxSampleValue"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                342 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TransferRange"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                346 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Indexed"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                351 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OPIProxy"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                515 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGRestartInterval"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                517 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGLosslessPredictors"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                518 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGPointTransforms"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                530 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "YCbCrSubSampling"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                531 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "YCbCrPositioning"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                18246 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Rating"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                18249 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "RatingPercent"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                33421 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CFARepeatPatternDim"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                34850 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ExposureProgram"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                34850 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ExposureProgram"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                34855 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ISOSpeedRatings"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                34855 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ISOSpeedRatings"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                34857 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Interlace"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                34859 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SelfTimerMode"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                34864 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SensitivityType"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37383 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "MeteringMode"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37383 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "MeteringMode"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37384 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "LightSource"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37384 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "LightSource"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37385 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Flash"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37385 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Flash"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37392 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "FocalPlaneResolutionUnit"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37396 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SubjectLocation"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37396 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SubjectArea"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37399 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SensingMethod"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                40961 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ColorSpace"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41488 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FocalPlaneResolutionUnit"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41492 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SubjectLocation"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41495 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SensingMethod"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41985 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "CustomRendered"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41986 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ExposureMode"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41987 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "WhiteBalance"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41989 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FocalLengthIn35mmFilm"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41990 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SceneCaptureType"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41991 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "GainControl"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41992 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Contrast"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41993 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Saturation"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41994 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Sharpness"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                41996 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SubjectDistanceRange"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                42080 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "CompositeImage"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                42081 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SourceImageNumberOfCompositeImage"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                50711 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CFALayout"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                50712 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "LinearizationTable"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                50713 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BlackLevelRepeatDim"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                50728 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "AsShotNeutral"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                50741 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "MakerNoteSafety"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                50778 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CalibrationIlluminant1"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                50779 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CalibrationIlluminant2"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                50879 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ColorimetricReference"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                51177 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DepthFormat"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                51180 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DepthUnits"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                51181 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DepthMeasureType"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                52529 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CalibrationIlluminant3"; Type = "Short"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                37377 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ShutterSpeedValue"; Type = "SRational"; Value = $DecimalValue } }
                37377 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ShutterSpeedValue"; Type = "SRational"; Value = $DecimalValue } }
                37379 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BrightnessValue"; Type = "SRational"; Value = $DecimalValue } }
                37379 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "BrightnessValue"; Type = "SRational"; Value = $DecimalValue } }
                37380 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ExposureBiasValue"; Type = "SRational"; Value = $DecimalValue } }
                37380 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ExposureBiasValue"; Type = "SRational"; Value = $DecimalValue } }
                37382 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SubjectDistance"; Type = "SRational"; Value = $DecimalValue } }
                37888 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "Temperature"; Type = "SRational"; Value = $DecimalValue } }
                37891 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "WaterDepth"; Type = "SRational"; Value = $DecimalValue } }
                37893 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "CameraElevationAngle"; Type = "SRational"; Value = $DecimalValue } }
                50715 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BlackLevelDeltaH"; Type = "SRational"; Value = $DecimalValue } }
                50716 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BlackLevelDeltaV"; Type = "SRational"; Value = $DecimalValue } }
                50721 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ColorMatrix1"; Type = "SRational"; Value = $DecimalValue } }
                50722 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ColorMatrix2"; Type = "SRational"; Value = $DecimalValue } }
                50723 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CameraCalibration1"; Type = "SRational"; Value = $DecimalValue } }
                50724 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CameraCalibration2"; Type = "SRational"; Value = $DecimalValue } }
                50725 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ReductionMatrix1"; Type = "SRational"; Value = $DecimalValue } }
                50726 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ReductionMatrix2"; Type = "SRational"; Value = $DecimalValue } }
                50730 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BaselineExposure"; Type = "SRational"; Value = $DecimalValue } }
                50739 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ShadowScale"; Type = "SRational"; Value = $DecimalValue } }
                50832 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "AsShotPreProfileMatrix"; Type = "SRational"; Value = $DecimalValue } }
                50834 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CurrentPreProfileMatrix"; Type = "SRational"; Value = $DecimalValue } }
                50964 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ForwardMatrix1"; Type = "SRational"; Value = $DecimalValue } }
                50965 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ForwardMatrix2"; Type = "SRational"; Value = $DecimalValue } }
                51044 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "FrameRate"; Type = "SRational"; Value = $DecimalValue } }
                51058 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TStop"; Type = "SRational"; Value = $DecimalValue } }
                51109 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "BaselineExposureOffset"; Type = "SRational"; Value = $DecimalValue } }
                52530 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CameraCalibration3"; Type = "SRational"; Value = $DecimalValue } }
                52531 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ColorMatrix3"; Type = "SRational"; Value = $DecimalValue } }
                52532 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ForwardMatrix3"; Type = "SRational"; Value = $DecimalValue } }
                52538 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ReductionMatrix3"; Type = "SRational"; Value = $DecimalValue } }
                344 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "XClipPathUnits"; Type = "SShort"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                345 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "YClipPathUnits"; Type = "SShort"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                28722 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "VignettingCorrParams"; Type = "SShort"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                28725 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ChromaticAberrationCorrParams"; Type = "SShort"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                28727 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "DistortionCorrParams"; Type = "SShort"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                34858 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "TimeZoneOffset"; Type = "SShort"; Value = [System.BitConverter]::ToUInt16($GetBitmap.Value, 0) } }
                #2 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Iop"; Name = "InteroperabilityVersion"; Type = "Undefined"; Value = [System.Text.Encoding]::ASCII.GetString($GetBitmap.Value) -replace "`0$"}}
                347 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "JPEGTables"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                34675 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "InterColorProfile"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                34856 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OECF"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                34856 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "OECF"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                36864 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ExifVersion"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                37121 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "ComponentsConfiguration"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                37388 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "SpatialFrequencyResponse"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                37389 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "Noise"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                37500 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "MakerNote"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                40960 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FlashpixVersion"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                41484 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SpatialFrequencyResponse"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                41728 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "FileSource"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                41729 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SceneType"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                41730 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "CFAPattern"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                41995 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "DeviceSettingDescription"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                42082 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Photo"; Name = "SourceExposureTimesOfCompositeImage"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                45057 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "MpfInfo"; Name = "MPFNumberOfImages"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                50341 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "PrintImageMatching"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                50828 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OriginalRawFileData"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                50831 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "AsShotICCProfile"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                50833 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "CurrentICCProfile"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                50972 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "RawImageDigest"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                50973 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OriginalRawFileDigest"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                51008 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OpcodeList1"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                51009 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OpcodeList2"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                51022 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "OpcodeList3"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52525 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileGainTableMap"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52533 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "IlluminantData1"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52534 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "IlluminantData2"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52535 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "IlluminantData3"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52543 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "RGBTables"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52544 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileGainTableMap2"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52548 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageSequenceInfo"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52550 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ImageStats"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                52551 { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Image"; Name = "ProfileDynamicRange"; Type = "Undefined"; Value = [string]$GetBitmap.Value } }
                # default {$ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Unknown"; Name = "Unknown"; Type = $GetBitmap.Value.GetType().FullName; Value = [string]$GetBitmap.Value}}
                default { $ExifResults += [PSCustomObject]@{ Source = "Exif"; Filename = $A.Name; Path = $A.DirectoryName; ID = $i; Group = "Unknown"; Name = "Unknown"; Type = $GetBitmap.Value.GetType().FullName; Value = [System.BitConverter]::ToInt32($GetBitmap.Value, 0) } }
            }
            return $ExifResults
        } }).Invoke() 

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Everything you want to know about your files"
$Form.Size = New-Object System.Drawing.Size(1024, 768)
# $Form.TopMost = $true
$Form.Font = New-Object System.Drawing.Font("Times New Roman", 12, [System.Drawing.FontStyle]::Regular)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedSingle"

$ButtonFile = New-Object System.Windows.Forms.Button
$ButtonFile.Text = "Select File(s)"
$ButtonFile.Location = New-Object System.Drawing.Point(20, 20)
$ButtonFile.Size = New-Object System.Drawing.Size(120, 40)

$ButtonFolder = New-Object System.Windows.Forms.Button
$ButtonFolder.Text = "Select Folder"
$ButtonFolder.Location = New-Object System.Drawing.Point(20, 65)
$ButtonFolder.Size = New-Object System.Drawing.Size(120, 40)

$ButtonReset = New-Object System.Windows.Forms.Button
$ButtonReset.Text = "Reset"
$ButtonReset.Location = New-Object System.Drawing.Point(20, 110)
$ButtonReset.Size = New-Object System.Drawing.Size(120, 40)


$ButtonQuit = New-Object System.Windows.Forms.Button
$ButtonQuit.Text = "Quit"
$ButtonQuit.Location = New-Object System.Drawing.Point(20, 155)
$ButtonQuit.Size = New-Object System.Drawing.Size(120, 40)

# $StatusTextBox = New-Object System.Windows.Forms.TextBox
$StatusTextBox = New-Object System.Windows.Forms.RichTextBox
$StatusTextBox.Multiline = $true
$StatusTextBox.ScrollBars = "Vertical"
$StatusTextBox.WordWrap = $true
$StatusTextBox.Size = New-Object System.Drawing.Size(390, 175)
$StatusTextBox.Location = New-Object System.Drawing.Point(450, 20)
$StatusTextBox.ReadOnly = $true
$StatusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$ListBoxSelectedExt = New-Object System.Windows.Forms.ListBox
$ListBoxSelectedExt.Size = New-Object System.Drawing.Size(140, 175)
$ListBoxSelectedExt.Location = New-Object System.Drawing.Point(855, 20)
$ListBoxSelectedExt.MultiColumn = $true

# Label for the extension filter buttons
$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Text = "Filter Extensions by First Letter:"
$filterLabel.Location = New-Object System.Drawing.Point(15, 250)
$filterLabel.AutoSize = $true
$filterLabel.Font = New-Object System.Drawing.Font($Form.Font, [System.Drawing.FontStyle]::Bold)

$panel = New-Object System.Windows.Forms.Panel
$panel.Size = New-Object System.Drawing.Size(980, 90)
$panel.Location = New-Object System.Drawing.Point(15, 270)
$panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle # Optional: to visualize the panel

$flowLayoutPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill # Make it fill the panel
$flowLayoutPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight # Example flow direction
$flowLayoutPanel.WrapContents = $true # Allow content wrapping
$flowLayoutPanel.BackColor = "LightBlue" # Optional: to visualize the flow layout panel

$ListViewExtensions = New-Object System.Windows.Forms.ListView
$ListViewExtensions.Size = New-Object System.Drawing.Size(980, 360)
$ListViewExtensions.Location = New-Object System.Drawing.Point(15, 370)
# Add columns (replace with your CSV headers)
$ListViewExtensions.Columns.Add("Ext", 100) | Out-Null
$ListViewExtensions.Columns.Add("Description", 300) | Out-Null
$ListViewExtensions.Columns.Add("Used by", 600) | Out-Null
$ListViewExtensions.View = [System.Windows.Forms.View]::Details
$ListViewExtensions.FullRowSelect = $true
$ListViewExtensions.GridLines = $true
$ListViewExtensions.CheckBoxes = $true
$ListViewExtensions.Name = "ListView0-9"

$FindOnly = New-Object System.Windows.Forms.CheckBox
$FindOnly.Text = "Find Only - Do Not Process"
$FindOnly.Location = New-Object System.Drawing.Point(150, 20)
$FindOnly.Checked = $false # Set initial state (optional)
$FindOnly.AutoSize = $true

$ShowInaccessable = New-Object System.Windows.Forms.CheckBox
$ShowInaccessable.Text = "Show Inaccessable Folders"
$ShowInaccessable.Location = New-Object System.Drawing.Point(150, 60)
$ShowInaccessable.Checked = $false # Set initial state (optional)
$ShowInaccessable.AutoSize = $true

$ButtonFile.Add_Click({
        $Results = Get-Files
        if (-not $Results[0]) { return } # Exit if user cancelled

        $PowerShell.Commands.Clear() # Clear previous commands
        $PowerShell.AddCommand('Get-AllData')
        $PowerShell.AddParameter("Path", $Results[0])
        $PowerShell.AddParameter("Include", $Results[1])
        $PowerShell.AddParameter("StatusTextBox", $StatusTextBox)
        $PowerShell.AddParameter("SyncHash", $SyncHash)
        $PowerShell.AddParameter("ListBoxSelectedExt", $ListBoxSelectedExt)
        $PowerShell.AddParameter("FindOnly", $FindOnly)
        $PowerShell.AddParameter("ShowInaccessable", $ShowInaccessable)
        $scanResults = $PowerShell.Invoke()

        if ($scanResults) {
            if ($scanResults[0].ContainsKey('FindOnlyResults')) {
                Show-FindOnlyResultsForm -ResultsData $scanResults[0] -ShowInaccessableCheckbox $ShowInaccessable
            } else {
                Show-ResultsForm -ResultsData $scanResults[0] -ShowInaccessableCheckbox $ShowInaccessable
            }
            $StatusTextBox.AppendText("Scan results are ready.`r`n")
        }
    })

$ButtonFolder.Add_Click({
        $Results = Get-Folder
        if (-not $Results[0]) { return } # Exit if user cancelled
        $PowerShell.Commands.Clear() # Clear previous commands
        $PowerShell.AddCommand('Get-AllData')
        $PowerShell.AddParameter("Path", $Results[0])
        $PowerShell.AddParameter("Include", $Results[1])
        $PowerShell.AddParameter("StatusTextBox", $StatusTextBox)
        $PowerShell.AddParameter("SyncHash", $SyncHash)
        $PowerShell.AddParameter("ListBoxSelectedExt", $ListBoxSelectedExt)
        $PowerShell.AddParameter("FindOnly", $FindOnly)
        $PowerShell.AddParameter("ShowInaccessable", $ShowInaccessable)
        $scanResults = $PowerShell.Invoke()

        if ($scanResults) {
            if ($scanResults[0].ContainsKey('FindOnlyResults')) {
                Show-FindOnlyResultsForm -ResultsData $scanResults[0] -ShowInaccessableCheckbox $ShowInaccessable
            } else {
                Show-ResultsForm -ResultsData $scanResults[0] -ShowInaccessableCheckbox $ShowInaccessable
            }
            $StatusTextBox.AppendText("Scan results are ready.`r`n")
        }
    })

$ButtonReset.Add_Click({
        $ListBoxSelectedExt.Items.Clear()
        $StatusTextBox.Clear()
        foreach ($control in $Form.Controls)
        {
            if ($control -is [System.Windows.Forms.ListView])
            {
                $control.Visible = $false
            }
        }
        foreach ($Control in $Form.Controls)
        {
            # Check if the control is a ListView
            if ($Control -is [System.Windows.Forms.ListView])
            {
                # Loop through the checked items in the current ListView
                foreach ($CheckedItem in $Control.CheckedItems)
                {
                    $CheckedItem.Checked = $false
                }
            }
        }
        $FindOnly.Checked = $false
        $ShowInaccessable.Checked = $false
    })

$ButtonQuit.Add_Click({

        $PowerShell.Dispose()
        $Runspace.Close()
        $Runspace.Dispose()
        $Form.Close() | Out-Null
    })

# Loop to create buttons A through Z
# We'll use ASCII values for this
# 'A' corresponds to ASCII 65, 'Z' to 90

foreach ($asciiValue in 64..90)
{
    $Letter = [char]$asciiValue

    if($Letter -match "@")
    {
        $Letter = "09"
    }

    # Create a new button
    $Button = New-Object System.Windows.Forms.Button
    $ListViewExtensions = New-Object System.Windows.Forms.ListView

    # Set the button's properties: Name and Text

    if ($Letter -eq "09")
    {
        $Button.Text = "0-9"
    }
    else
    {
        $Button.Text = $Letter
    }
    $Button.Name = "Button" + $Letter
    $Button.Text = $Letter

    # Optional: Customize button size and location
    # If using FlowLayoutPanel, location and size might be managed automatically

    # Optional: Add an event handler for button clicks (e.g., display a message box)
    $Button.Add_Click({
            foreach ($control in $Form.Controls)
            {
                if ($control -is [System.Windows.Forms.ListView])
                {
                    $control.Visible = $false
                }
            }
            $CurrentLetter = ($this.Text)
            $SelectedListView = $form.Controls["ListView$($CurrentLetter)"]
            $SelectedListView.Visible = $true
        })

    # Add the button to the FlowLayoutPanel
    $flowLayoutPanel.Controls.Add($Button)
    
    $ListViewExtensions.Size = New-Object System.Drawing.Size(980, 360)
    $ListViewExtensions.Location = New-Object System.Drawing.Point(15, 370)
    # Add columns (replace with your CSV headers)
    $ListViewExtensions.Columns.Add("Ext", 100) | Out-Null
    $ListViewExtensions.Columns.Add("Description", 300) | Out-Null
    $ListViewExtensions.Columns.Add("Used by", 600) | Out-Null
    $ListViewExtensions.View = [System.Windows.Forms.View]::Details
    $ListViewExtensions.FullRowSelect = $true
    $ListViewExtensions.GridLines = $true
    $ListViewExtensions.CheckBoxes = $true
    $ListViewExtensions.Name = "ListView$($Letter)"
    $Form.Controls.Add($ListViewExtensions)    
    $ListViewExtensions.Add_ItemCheck({
            param($sender, $e)

            $listView = $sender
            $itemText = $listView.Items[$e.Index].Text
            if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
            {
                # $e.Index gives the index of the item being checked/unchecked
                # $e.CurrentValue gives the check state before the change

                # $itemText = $listView.Items[$e.Index].Text
                # $oldState = $e.CurrentValue
                # $newState = $e.NewValue # The new state of the checkbox

                # Write-Host "Item '$itemText' changed from '$oldState' to '$newState'"
                # You can add your custom logic here based on the checkbox state change

                if ($e.NewValue -eq "Checked")
                {
                    $SyncHash.CheckedExtensions.Add($itemText, $itemText)
                }

                if($null -ne $SyncHash.CheckedExtensions)
                {
                    $ListBoxSelectedExt.Items.Clear()
                    $SyncHash.CheckedExtensions.Keys | ForEach-Object { $ListBoxSelectedExt.Items.Add($_) }
                }
            }
            elseif($e.NewValue -eq [System.Windows.Forms.CheckState]::Unchecked)
            {
                $SyncHash.CheckedExtensions.Remove($($itemText))
                if($null -ne $SyncHash.CheckedExtensions)
                {
                    $ListBoxSelectedExt.Items.Clear()
                    $SyncHash.CheckedExtensions.Keys | ForEach-Object { $ListBoxSelectedExt.Items.Add($_) }
                }
            
            }
        })


    Get-FilteredExtensions -CurrentLetter $Letter -CsvData $CsvData | Out-Null
}

foreach ($control in $Form.Controls)
{
    if ($control -is [System.Windows.Forms.ListView])
    {
        $control.Visible = $false
    }
}

$Form.Controls.Add($ButtonFile)
$Form.Controls.Add($ButtonFolder)
$Form.Controls.Add($ButtonReset)
$Form.Controls.Add($ButtonQuit)
$Form.Controls.Add($StatusTextBox)
$Form.Controls.Add($ListBoxSelectedExt)
$Form.Controls.Add($panel) # Add the panel to the form
$Form.Controls.Add($filterLabel)
$panel.Controls.Add($flowLayoutPanel)
$Form.Controls.Add($FindOnly)
$Form.Controls.Add($ShowInaccessable)
$Form.ShowDialog()
