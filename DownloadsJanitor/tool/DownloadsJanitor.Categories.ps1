# DownloadsJanitor.Categories.ps1
# ---------------------------------------------------------------------------
# The ONLY file you need to edit to tune how files are sorted.
# It is dot-sourced by the engine, so a mistake here cannot break the engine.
#
# Returns a hashtable with: Map, SkipNames, SkipSuffixes, NeverMoveExt,
#                           RecordsCategories, LegacyFolders
#
# Extension keys must be lowercase and start with a dot.
# Category values become folder names, so they must be legal folder names.
# ---------------------------------------------------------------------------

$Map = @{
    # ---------------------------------------------------------- Artwork & Design
    '.psd'='Artwork & Design'; '.psb'='Artwork & Design'; '.psdc'='Artwork & Design'
    '.pdd'='Artwork & Design'; '.procreate'='Artwork & Design'; '.clip'='Artwork & Design'
    '.lip'='Artwork & Design'; '.kra'='Artwork & Design'; '.krz'='Artwork & Design'
    '.sai'='Artwork & Design'; '.sai2'='Artwork & Design'; '.reb'='Artwork & Design'
    '.afphoto'='Artwork & Design'; '.afdesign'='Artwork & Design'; '.afpub'='Artwork & Design'
    '.xcf'='Artwork & Design'; '.rif'='Artwork & Design'; '.riff'='Artwork & Design'
    '.mdp'='Artwork & Design'; '.pdn'='Artwork & Design'; '.psp'='Artwork & Design'
    '.pspimage'='Artwork & Design'; '.cpt'='Artwork & Design'; '.tvpp'='Artwork & Design'
    '.ora'='Artwork & Design'; '.skb'='Artwork & Design'; '.ai'='Artwork & Design'
    '.svg'='Artwork & Design'; '.svgz'='Artwork & Design'; '.eps'='Artwork & Design'
    '.ps'='Artwork & Design'; '.cdr'='Artwork & Design'; '.cdt'='Artwork & Design'
    '.cmx'='Artwork & Design'; '.indd'='Artwork & Design'; '.indl'='Artwork & Design'
    '.indt'='Artwork & Design'; '.indb'='Artwork & Design'; '.idml'='Artwork & Design'
    '.inx'='Artwork & Design'; '.fig'='Artwork & Design'; '.sketch'='Artwork & Design'
    '.xd'='Artwork & Design'; '.qxp'='Artwork & Design'; '.qxd'='Artwork & Design'
    '.sla'='Artwork & Design'; '.scd'='Artwork & Design'; '.pub'='Artwork & Design'
    '.xar'='Artwork & Design'; '.pxd'='Artwork & Design'; '.pxm'='Artwork & Design'
    '.emf'='Artwork & Design'; '.wmf'='Artwork & Design'; '.odg'='Artwork & Design'

    # ------------------------------------------------------------------- Photos
    '.jpg'='Photos'; '.jpeg'='Photos'; '.jpe'='Photos'; '.jfif'='Photos'
    '.png'='Photos'; '.apng'='Photos'; '.gif'='Photos'; '.webp'='Photos'
    '.avif'='Photos'; '.jxl'='Photos'; '.heic'='Photos'; '.heif'='Photos'
    '.bmp'='Photos'; '.dib'='Photos'; '.tif'='Photos'; '.tiff'='Photos'
    '.ico'='Photos'; '.cur'='Photos'; '.tga'='Photos'; '.jp2'='Photos'
    '.j2k'='Photos'; '.jpf'='Photos'; '.jpx'='Photos'; '.ppm'='Photos'
    '.pgm'='Photos'; '.pbm'='Photos'; '.pnm'='Photos'

    # --------------------------------------------------------------- Camera Raw
    '.cr2'='Camera Raw'; '.cr3'='Camera Raw'; '.crw'='Camera Raw'; '.nef'='Camera Raw'
    '.nrw'='Camera Raw'; '.arw'='Camera Raw'; '.srf'='Camera Raw'; '.sr2'='Camera Raw'
    '.raf'='Camera Raw'; '.orf'='Camera Raw'; '.ori'='Camera Raw'; '.rw2'='Camera Raw'
    '.rwl'='Camera Raw'; '.raw'='Camera Raw'; '.dng'='Camera Raw'; '.pef'='Camera Raw'
    '.srw'='Camera Raw'; '.x3f'='Camera Raw'; '.3fr'='Camera Raw'; '.fff'='Camera Raw'
    '.iiq'='Camera Raw'; '.mos'='Camera Raw'; '.mrw'='Camera Raw'; '.erf'='Camera Raw'
    '.dcr'='Camera Raw'; '.kdc'='Camera Raw'; '.k25'='Camera Raw'; '.mef'='Camera Raw'
    '.gpr'='Camera Raw'
    '.xmp'='Camera Raw'; '.aae'='Camera Raw'; '.thm'='Camera Raw'; '.pp3'='Camera Raw'
    '.dop'='Camera Raw'; '.on1'='Camera Raw'; '.cos'='Camera Raw'; '.cot'='Camera Raw'
    '.lrcat'='Camera Raw'; '.lrdata'='Camera Raw'; '.lrprev'='Camera Raw'
    '.cocatalogdb'='Camera Raw'; '.cosessiondb'='Camera Raw'

    # ------------------------------------------------------------------ 3D & CAD
    '.blend'='3D & CAD'; '.blend1'='3D & CAD'; '.blend2'='3D & CAD'; '.blend3'='3D & CAD'
    '.ma'='3D & CAD'; '.mb'='3D & CAD'; '.max'='3D & CAD'; '.c4d'='3D & CAD'
    '.ztl'='3D & CAD'; '.zpr'='3D & CAD'; '.zsc'='3D & CAD'; '.sbs'='3D & CAD'
    '.sbsar'='3D & CAD'; '.spp'='3D & CAD'; '.spsm'='3D & CAD'; '.hip'='3D & CAD'
    '.hipnc'='3D & CAD'; '.hiplc'='3D & CAD'; '.bgeo'='3D & CAD'; '.lxo'='3D & CAD'
    '.lxf'='3D & CAD'; '.lwo'='3D & CAD'; '.lws'='3D & CAD'; '.3dm'='3D & CAD'
    '.skp'='3D & CAD'; '.zprj'='3D & CAD'; '.zpac'='3D & CAD'; '.tbscene'='3D & CAD'
    '.tbmat'='3D & CAD'; '.unitypackage'='3D & CAD'; '.prefab'='3D & CAD'
    '.unity'='3D & CAD'; '.uasset'='3D & CAD'; '.umap'='3D & CAD'; '.uproject'='3D & CAD'
    '.upk'='3D & CAD'
    '.obj'='3D & CAD'; '.mtl'='3D & CAD'; '.fbx'='3D & CAD'; '.dae'='3D & CAD'
    '.stl'='3D & CAD'; '.ply'='3D & CAD'; '.3ds'='3D & CAD'; '.glb'='3D & CAD'
    '.gltf'='3D & CAD'; '.usd'='3D & CAD'; '.usda'='3D & CAD'; '.usdc'='3D & CAD'
    '.usdz'='3D & CAD'; '.abc'='3D & CAD'; '.x3d'='3D & CAD'; '.wrl'='3D & CAD'
    '.vox'='3D & CAD'; '.bvh'='3D & CAD'
    '.exr'='3D & CAD'; '.hdr'='3D & CAD'; '.dds'='3D & CAD'; '.ktx'='3D & CAD'
    '.ktx2'='3D & CAD'
    '.step'='3D & CAD'; '.stp'='3D & CAD'; '.iges'='3D & CAD'; '.igs'='3D & CAD'
    '.dwg'='3D & CAD'; '.dxf'='3D & CAD'; '.dwf'='3D & CAD'; '.dwfx'='3D & CAD'
    '.3mf'='3D & CAD'; '.amf'='3D & CAD'; '.f3d'='3D & CAD'; '.rvt'='3D & CAD'
    '.rfa'='3D & CAD'; '.rte'='3D & CAD'; '.sldprt'='3D & CAD'; '.sldasm'='3D & CAD'
    '.slddrw'='3D & CAD'; '.ipt'='3D & CAD'; '.iam'='3D & CAD'; '.idw'='3D & CAD'
    '.catpart'='3D & CAD'; '.catproduct'='3D & CAD'; '.prt'='3D & CAD'; '.asm'='3D & CAD'
    '.x_t'='3D & CAD'; '.x_b'='3D & CAD'; '.sat'='3D & CAD'
    '.gcode'='3D & CAD'; '.nc'='3D & CAD'; '.cnc'='3D & CAD'; '.dst'='3D & CAD'
    '.emb'='3D & CAD'; '.plt'='3D & CAD'

    # -------------------------------------------------------------------- Video
    '.mp4'='Video'; '.m4v'='Video'; '.mkv'='Video'; '.avi'='Video'; '.mov'='Video'
    '.wmv'='Video'; '.flv'='Video'; '.f4v'='Video'; '.webm'='Video'; '.mpg'='Video'
    '.mpeg'='Video'; '.m2v'='Video'; '.mpv'='Video'; '.mts'='Video'; '.m2ts'='Video'
    '.ts'='Video'; '.vob'='Video'; '.3gp'='Video'; '.3g2'='Video'; '.ogv'='Video'
    '.rm'='Video'; '.rmvb'='Video'; '.asf'='Video'; '.divx'='Video'; '.qt'='Video'
    '.dv'='Video'; '.mxf'='Video'
    '.srt'='Video'; '.vtt'='Video'; '.ass'='Video'; '.ssa'='Video'; '.sub'='Video'
    '.idx'='Video'
    '.braw'='Video'; '.r3d'='Video'; '.ari'='Video'; '.cine'='Video'
    '.aep'='Video'; '.aepx'='Video'; '.aet'='Video'; '.prproj'='Video'; '.prel'='Video'
    '.ppj'='Video'; '.drp'='Video'; '.drt'='Video'; '.drfx'='Video'; '.fcpxml'='Video'
    '.fcpbundle'='Video'; '.motn'='Video'; '.avp'='Video'; '.avb'='Video'; '.veg'='Video'
    '.camproj'='Video'; '.tscproj'='Video'; '.trec'='Video'; '.nk'='Video'; '.nknc'='Video'
    '.comp'='Video'; '.tnz'='Video'; '.xstage'='Video'; '.fla'='Video'; '.xfl'='Video'
    '.swf'='Video'; '.spine'='Video'; '.cmo3'='Video'; '.can3'='Video'; '.moc3'='Video'

    # ------------------------------------------------------------ Music & Audio
    '.mp3'='Music & Audio'; '.wav'='Music & Audio'; '.flac'='Music & Audio'
    '.aac'='Music & Audio'; '.ogg'='Music & Audio'; '.oga'='Music & Audio'
    '.m4a'='Music & Audio'; '.wma'='Music & Audio'; '.opus'='Music & Audio'
    '.aiff'='Music & Audio'; '.aif'='Music & Audio'; '.aifc'='Music & Audio'
    '.alac'='Music & Audio'; '.ape'='Music & Audio'; '.wv'='Music & Audio'
    '.mka'='Music & Audio'; '.amr'='Music & Audio'; '.au'='Music & Audio'
    '.caf'='Music & Audio'; '.dsf'='Music & Audio'; '.dff'='Music & Audio'
    '.mid'='Music & Audio'; '.midi'='Music & Audio'; '.m3u'='Music & Audio'
    '.m3u8'='Music & Audio'; '.pls'='Music & Audio'
    '.als'='Music & Audio'; '.alp'='Music & Audio'; '.adg'='Music & Audio'
    '.adv'='Music & Audio'; '.agr'='Music & Audio'; '.ams'='Music & Audio'
    '.amxd'='Music & Audio'; '.asd'='Music & Audio'; '.alc'='Music & Audio'
    '.ask'='Music & Audio'; '.flp'='Music & Audio'; '.fst'='Music & Audio'
    '.fsc'='Music & Audio'; '.logicx'='Music & Audio'; '.logic'='Music & Audio'
    '.band'='Music & Audio'; '.rpp'='Music & Audio'; '.rpl'='Music & Audio'
    '.rfxchain'='Music & Audio'; '.reapeaks'='Music & Audio'; '.ptx'='Music & Audio'
    '.pts'='Music & Audio'; '.ptf'='Music & Audio'; '.ptt'='Music & Audio'
    '.cpr'='Music & Audio'; '.npr'='Music & Audio'; '.song'='Music & Audio'
    '.bwproject'='Music & Audio'; '.bwpreset'='Music & Audio'; '.bwclip'='Music & Audio'
    '.reason'='Music & Audio'; '.rns'='Music & Audio'; '.rps'='Music & Audio'
    '.cwp'='Music & Audio'; '.cwb'='Music & Audio'; '.aup'='Music & Audio'
    '.aup3'='Music & Audio'; '.xrns'='Music & Audio'; '.xrni'='Music & Audio'
    '.sf2'='Music & Audio'; '.sfz'='Music & Audio'; '.nki'='Music & Audio'
    '.nkm'='Music & Audio'; '.nkx'='Music & Audio'; '.nkc'='Music & Audio'
    '.nks'='Music & Audio'; '.exs'='Music & Audio'; '.rex'='Music & Audio'
    '.rx2'='Music & Audio'; '.fxp'='Music & Audio'; '.fxb'='Music & Audio'
    '.vstpreset'='Music & Audio'; '.aupreset'='Music & Audio'; '.h2p'='Music & Audio'
    '.nmsv'='Music & Audio'; '.vital'='Music & Audio'
    '.mscz'='Music & Audio'; '.mscx'='Music & Audio'; '.sib'='Music & Audio'
    '.mus'='Music & Audio'; '.musx'='Music & Audio'; '.musicxml'='Music & Audio'
    '.mxl'='Music & Audio'; '.gp'='Music & Audio'; '.gp3'='Music & Audio'
    '.gp4'='Music & Audio'; '.gp5'='Music & Audio'; '.gpx'='Music & Audio'
    '.gp7'='Music & Audio'

    # -------------------------------------------------------------------- Fonts
    '.ttf'='Fonts'; '.otf'='Fonts'; '.ttc'='Fonts'; '.otc'='Fonts'
    '.woff'='Fonts'; '.woff2'='Fonts'; '.eot'='Fonts'; '.pfb'='Fonts'
    '.pfa'='Fonts'; '.pfm'='Fonts'; '.afm'='Fonts'; '.fon'='Fonts'
    '.fnt'='Fonts'; '.dfont'='Fonts'; '.bdf'='Fonts'; '.pcf'='Fonts'
    '.sfd'='Fonts'; '.glyphs'='Fonts'; '.vfb'='Fonts'; '.vfc'='Fonts'
    '.vfj'='Fonts'; '.designspace'='Fonts'; '.fea'='Fonts'; '.ttx'='Fonts'

    # ------------------------------------------------------- Brushes & Presets
    '.abr'='Brushes & Presets'; '.tpl'='Brushes & Presets'; '.pat'='Brushes & Presets'
    '.grd'='Brushes & Presets'; '.asl'='Brushes & Presets'; '.aco'='Brushes & Presets'
    '.ase'='Brushes & Presets'; '.acb'='Brushes & Presets'; '.acv'='Brushes & Presets'
    '.csh'='Brushes & Presets'; '.atn'='Brushes & Presets'; '.8bf'='Brushes & Presets'
    '.brushset'='Brushes & Presets'; '.brush'='Brushes & Presets'
    '.swatches'='Brushes & Presets'; '.sut'='Brushes & Presets'
    '.kpp'='Brushes & Presets'; '.bundle'='Brushes & Presets'; '.myb'='Brushes & Presets'
    '.gbr'='Brushes & Presets'; '.gih'='Brushes & Presets'; '.vbr'='Brushes & Presets'
    '.gpl'='Brushes & Presets'; '.afbrushes'='Brushes & Presets'
    '.afmacros'='Brushes & Presets'; '.afstyles'='Brushes & Presets'
    '.afpalette'='Brushes & Presets'; '.aftemplate'='Brushes & Presets'
    '.lrtemplate'='Brushes & Presets'; '.dcp'='Brushes & Presets'
    '.costyle'='Brushes & Presets'; '.icc'='Brushes & Presets'; '.icm'='Brushes & Presets'
    '.cube'='Brushes & Presets'; '.3dl'='Brushes & Presets'; '.look'='Brushes & Presets'
    '.lut'='Brushes & Presets'; '.clr'='Brushes & Presets'; '.soc'='Brushes & Presets'

    # ---------------------------------------------------------------- Documents
    '.pdf'='Documents'; '.doc'='Documents'; '.docx'='Documents'; '.docm'='Documents'
    '.dot'='Documents'; '.dotx'='Documents'; '.dotm'='Documents'; '.odt'='Documents'
    '.ott'='Documents'; '.fodt'='Documents'; '.rtf'='Documents'; '.txt'='Documents'
    '.md'='Documents'; '.markdown'='Documents'; '.epub'='Documents'; '.mobi'='Documents'
    '.azw'='Documents'; '.azw3'='Documents'; '.djvu'='Documents'; '.fb2'='Documents'
    '.tex'='Documents'; '.pages'='Documents'; '.wpd'='Documents'; '.wps'='Documents'
    '.xps'='Documents'; '.oxps'='Documents'; '.chm'='Documents'
    '.cbz'='Documents'; '.cbr'='Documents'; '.cb7'='Documents'; '.cbt'='Documents'
    '.one'='Documents'; '.onepkg'='Documents'; '.onetoc2'='Documents'
    '.vsd'='Documents'; '.vsdx'='Documents'; '.vsdm'='Documents'; '.vss'='Documents'
    '.vsx'='Documents'; '.vst'='Documents'; '.vstx'='Documents'
    '.mpp'='Documents'; '.mpx'='Documents'
    '.html'='Documents'; '.htm'='Documents'
    '.ppt'='Documents'; '.pptx'='Documents'; '.pptm'='Documents'; '.pps'='Documents'
    '.ppsx'='Documents'; '.ppsm'='Documents'; '.pot'='Documents'; '.potx'='Documents'
    '.potm'='Documents'; '.odp'='Documents'; '.otp'='Documents'; '.kth'='Documents'
    '.msg'='Documents'; '.eml'='Documents'; '.emlx'='Documents'; '.mbox'='Documents'
    '.mbx'='Documents'; '.pst'='Documents'; '.olm'='Documents'; '.nsf'='Documents'
    '.oft'='Documents'
    '.ics'='Documents'; '.ical'='Documents'; '.ifb'='Documents'; '.vcs'='Documents'
    '.vcf'='Documents'; '.vcard'='Documents'
    '.p7m'='Documents'; '.p7s'='Documents'; '.asice'='Documents'; '.sce'='Documents'
    '.bdoc'='Documents'; '.edoc'='Documents'; '.ddoc'='Documents'; '.xsig'='Documents'
    '.asc'='Documents'; '.sig'='Documents'

    # ------------------------------------------------------------- Spreadsheets
    '.xls'='Spreadsheets'; '.xlsx'='Spreadsheets'; '.xlsm'='Spreadsheets'
    '.xlsb'='Spreadsheets'; '.xlt'='Spreadsheets'; '.xltx'='Spreadsheets'
    '.xltm'='Spreadsheets'; '.ods'='Spreadsheets'; '.ots'='Spreadsheets'
    '.fods'='Spreadsheets'; '.csv'='Spreadsheets'; '.tsv'='Spreadsheets'
    '.numbers'='Spreadsheets'; '.dif'='Spreadsheets'; '.slk'='Spreadsheets'
    '.prn'='Spreadsheets'; '.dbf'='Spreadsheets'

    # --------------------------------------------------------------- Accounting
    '.ofx'='Accounting'; '.qfx'='Accounting'; '.qbo'='Accounting'; '.qif'='Accounting'
    '.qbx'='Accounting'; '.qby'='Accounting'; '.qbw'='Accounting'; '.qbb'='Accounting'
    '.qbm'='Accounting'; '.qba'='Accounting'; '.qbr'='Accounting'; '.nd'='Accounting'
    '.tlg'='Accounting'; '.iif'='Accounting'; '.qdf'='Accounting'; '.qdfx'='Accounting'
    '.qel'='Accounting'; '.qph'='Accounting'; '.ptb'='Accounting'; '.saj'='Accounting'
    '.myo'='Accounting'; '.myox'='Accounting'; '.gnucash'='Accounting'; '.gcm'='Accounting'
    '.mny'='Accounting'; '.mbf'='Accounting'; '.moneydance'='Accounting'
    '.mt940'='Accounting'; '.sta'='Accounting'; '.camt'='Accounting'; '.aba'='Accounting'
    '.pain'='Accounting'; '.900'='Accounting'; '.tsf'='Accounting'; '.tcp'='Accounting'
    '.tax'='Accounting'

    # ------------------------------------------------------------- Certificates
    '.pfx'='Certificates'; '.p12'='Certificates'; '.cer'='Certificates'
    '.crt'='Certificates'; '.der'='Certificates'; '.pem'='Certificates'
    '.csr'='Certificates'; '.jks'='Certificates'; '.keystore'='Certificates'
    '.p7b'='Certificates'; '.spc'='Certificates'

    # ---------------------------------------------------------- Zips & Archives
    '.zip'='Zips & Archives'; '.zipx'='Zips & Archives'; '.rar'='Zips & Archives'
    '.7z'='Zips & Archives'; '.tar'='Zips & Archives'; '.gz'='Zips & Archives'
    '.tgz'='Zips & Archives'; '.bz2'='Zips & Archives'; '.tbz2'='Zips & Archives'
    '.xz'='Zips & Archives'; '.txz'='Zips & Archives'; '.zst'='Zips & Archives'
    '.lz'='Zips & Archives'; '.lzma'='Zips & Archives'; '.lz4'='Zips & Archives'
    '.cab'='Zips & Archives'; '.arj'='Zips & Archives'; '.ace'='Zips & Archives'
    '.z'='Zips & Archives'; '.lzh'='Zips & Archives'; '.iso'='Zips & Archives'

    # --------------------------------------------------------- Apps & Installers
    '.exe'='Apps & Installers'; '.msi'='Apps & Installers'; '.msix'='Apps & Installers'
    '.msixbundle'='Apps & Installers'; '.appx'='Apps & Installers'
    '.appxbundle'='Apps & Installers'; '.msu'='Apps & Installers'
    '.msp'='Apps & Installers'; '.apk'='Apps & Installers'; '.aab'='Apps & Installers'
    '.dmg'='Apps & Installers'; '.pkg'='Apps & Installers'; '.deb'='Apps & Installers'
    '.rpm'='Apps & Installers'; '.appimage'='Apps & Installers'

    # --------------------------------------------------------------------- Code
    '.ps1'='Code'; '.psm1'='Code'; '.psd1'='Code'; '.bat'='Code'; '.cmd'='Code'
    '.sh'='Code'; '.py'='Code'; '.ipynb'='Code'; '.js'='Code'; '.jsx'='Code'
    '.tsx'='Code'; '.java'='Code'; '.c'='Code'; '.cpp'='Code'; '.h'='Code'
    '.cs'='Code'; '.go'='Code'; '.rs'='Code'; '.rb'='Code'; '.php'='Code'
    '.sql'='Code'; '.css'='Code'; '.vbs'='Code'; '.vba'='Code'; '.yaml'='Code'
    '.yml'='Code'
}

# --- Files that are NEVER moved, matched by exact name (case-insensitive) ----
$SkipNames = @(
    'desktop.ini', 'thumbs.db', '.ds_store',
    '_what-is-this.txt', 'where are my files.html'
)

# --- NEVER moved, matched by how the FILE NAME ENDS (not the extension) ------
# Partial downloads: "movie.mkv.!qB" has extension ".!qb" but "big.iso.aria2"
# would otherwise be classified as an unknown extension. Suffix matching
# catches both forms.
$SkipSuffixes = @(
    '.crdownload', '.part', '.partial', '.tmp', '.download', '.opdownload',
    '.!qb', '.!ut', '.!bt', '.aria2', '.bc!', '.filepart', '.dctmp',
    '.crswap', '.td', '.temp', '.tmp~'
)

# --- NEVER moved, matched by extension --------------------------------------
# .lnk/.url/.website : someone deliberately put that pointer there.
# .ost               : live Outlook cache bound to a fixed path.
$NeverMoveExt = @('.lnk', '.url', '.website', '.ost')

# --- Categories exempt from the "old files" report ---------------------------
# An archived contract is doing its job. An installer nobody ran in two years
# is clutter. Only the clutter categories get reported.
$RecordsCategories = @(
    'Documents', 'Spreadsheets', 'Accounting', 'Certificates', '3D & CAD'
)

# --- Folder names from older versions of this tool ---------------------------
# NEVER remove a line here. These are recognised as ours so that a folder made
# by an older version is not reported to the user as junk they could delete.
$LegacyFolders = @(
    'Images', 'Audio', 'Presentations', 'Archives', 'Installers',
    'Data', 'Torrents', 'Shortcuts', 'Other', 'Everything Else'
)

@{
    Map               = $Map
    SkipNames         = $SkipNames
    SkipSuffixes      = $SkipSuffixes
    NeverMoveExt      = $NeverMoveExt
    RecordsCategories = $RecordsCategories
    LegacyFolders     = $LegacyFolders
}
