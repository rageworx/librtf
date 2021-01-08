#include <windows.h>

#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <string>

#include <fcntl.h>
#include <sys/stat.h>
#include <io.h>
#include <olectl.h>

#include "librtf.h"

using namespace std;

////////////////////////////////////////////////////////////////////////////////

// RTF document formatting params
static RTF_DOCUMENT_FORMAT rtfDocFormat = {0};

// RTF section formatting params
static RTF_SECTION_FORMAT rtfSecFormat = {0};

// RTF paragraph formatting params
static RTF_PARAGRAPH_FORMAT rtfParFormat = {0};

// RTF table row formatting params
static RTF_TABLEROW_FORMAT rtfRowFormat = {0};

// RTF table cell formatting params
static RTF_TABLECELL_FORMAT rtfCellFormat = {0};

// RTF library global params
static FILE*        rtfFile = NULL;
static string       rtfFontTable;
static string       rtfColorTable;
static IPicture*    rtfPicture = NULL;

static void strcats( char* ob, const char* ib, size_t obsz )
{
    if ( ( ob == NULL ) || ( ib == NULL ) || ( obsz == 0 ) )
        return;

    strcat_s( ob, obsz, ib );
}

// Creates new RTF document
RTF_ERROR_TYPE librtf::open( const char* filename, const char* fonts, const char* colors,
                             RTF_DOCUMENT_FORMAT* fmt )
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_SUCCESS;

    // Initialize global params
    librtf::init();

    // Set RTF document font table
    if ( fonts != NULL )
    {
        if ( strlen( fonts ) > 0 )
            librtf::set_fonttable(fonts);
    }

    // Set RTF document color table
    if ( colors != NULL )
    {
        if ( strlen( colors ) > 0 )
            librtf::set_colortable(colors);
    }

    // Set Document format
    if ( fmt != NULL )
    {
        librtf::set_documentformat( fmt );
    }

    // Create RTF document
    rtfFile = fopen( filename, "wb" );

    if ( rtfFile != NULL )
    {
        // Write RTF document header
        if ( librtf::write_header() == false )
        {
            fclose( rtfFile );
            rtfFile = NULL;
            error = RTF_HEADER_ERROR;
            return error;
        }

        // Write RTF document formatting properties
        if ( librtf::write_documentformat() == false )
        {
            fclose( rtfFile );
            rtfFile = NULL;
            error = RTF_DOCUMENTFORMAT_ERROR;
            return error;
        }

        // Create first RTF document section with default formatting
        if ( librtf::write_sectionformat() == false )
        {
            fclose( rtfFile );
            rtfFile = NULL;
            error = RTF_SECTIONFORMAT_ERROR;
            return error;
        }
    }
    else
    {
        error = RTF_OPEN_ERROR;
    }

    // Return error flag
    return error;
}

// Closes created RTF document
RTF_ERROR_TYPE librtf::close()
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_SUCCESS;

    if( rtfFile != NULL )
    {
        // Free IPicture object
        if ( rtfPicture != NULL )
        {
            rtfPicture->Release();
            rtfPicture = NULL;
        }

        // Write RTF document end part
        char rtfText[] = "\n\\par}";
        fwrite( rtfText, 1, strlen(rtfText), rtfFile );

        // Close RTF document
        if ( fclose(rtfFile) != 0 )
            error = RTF_CLOSE_ERROR;

        rtfFile = NULL;
    }

    if ( rtfParFormat.paragraphText  != NULL )
    {
        delete[] rtfParFormat.paragraphText;
        rtfParFormat.paragraphText = NULL;
    }

    // Return error flag
    return error;
}

// Writes RTF document header
bool librtf::write_header()
{
    // Set error flag
    bool result = true;

    // Standard RTF document header
    string wrbuff;

    if ( rtfFontTable.size() > 0 )
    {
        wrbuff += "{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl";
        wrbuff += rtfFontTable;
        wrbuff += "}";
    }

    if ( rtfColorTable.size() > 0 )
    {
        wrbuff += "{\\colortbl";
        wrbuff += rtfColorTable;
        wrbuff += "}";
    }

    wrbuff += "{\\*\\generator librtf v1.2;}\n";
    wrbuff += "{\\info{\\author none}{\\company none}}";

    // Writes standard RTF document header part
    if ( rtfFile != NULL )
    {
        if ( fwrite( wrbuff.c_str(), 1, wrbuff.size(), rtfFile ) < wrbuff.size() )
            result = false;
    }
    else
    {
        result = false;
    }

    wrbuff.clear();

    // Return error flag
    return result;
}

// Sets global RTF library params
void librtf::init()
{
    // Set RTF document default font table
    if ( rtfFontTable.size() > 0 )
        rtfFontTable.clear();

    rtfFontTable += "{\\f0\\froman\\fcharset0\\cpg1252 Times New Roman}";
    rtfFontTable += "{\\f1\\fswiss\\fcharset0\\cpg1252 Arial}";
    rtfFontTable += "{\\f2\\fmodern\\fcharset0\\cpg1252 Courier New}";
    rtfFontTable += "{\\f3\\fscript\\fcharset0\\cpg1252 Cursive}";
    rtfFontTable += "{\\f4\\fdecor\\fcharset0\\cpg1252 Old English}";
    rtfFontTable += "{\\f5\\ftech\\fcharset0\\cpg1252 Symbol}";
    rtfFontTable += "{\\f6\\fbidi\\fcharset0\\cpg1252 Miriam}";

    // Set RTF document default color table
    if ( rtfColorTable.size() > 0 )
        rtfColorTable.clear();

    rtfColorTable += "\\red0\\green0\\blue0;";
    rtfColorTable += "\\red255\\green0\\blue0;";
    rtfColorTable += "\\red0\\green255\\blue0;";
    rtfColorTable += "\\red0\\green0\\blue255;";
    rtfColorTable += "\\red255\\green255\\blue0;";
    rtfColorTable += "\\red255\\green0\\blue255;";
    rtfColorTable += "\\red0\\green255\\blue255;";
    rtfColorTable += "\\red255\\green255\\blue255;";
    rtfColorTable += "\\red128\\green0\\blue0;";
    rtfColorTable += "\\red0\\green128\\blue0;";
    rtfColorTable += "\\red0\\green0\\blue128;";
    rtfColorTable += "\\red128\\green128\\blue0;";
    rtfColorTable += "\\red128\\green0\\blue128;";
    rtfColorTable += "\\red0\\green128\\blue128;";
    rtfColorTable += "\\red128\\green128\\blue128;";

    // Set default formatting
    librtf::set_defaultformat();
}

// Sets default RTF document formatting
void librtf::set_defaultformat()
{
    // Set default RTF document formatting properties
    RTF_DOCUMENT_FORMAT df = { RTF_DOCUMENTVIEWKIND_PAGE,
                               100, 12240, 15840, 1800, 1800, 1440, 1440,
                               false, 0, false };

    librtf::set_documentformat( &df );

    // Set default RTF section formatting properties
    RTF_SECTION_FORMAT sf = { RTF_SECTIONBREAK_CONTINUOUS,
                              false, true,
                              12240, 15840, 1800, 1800, 1440, 1440, 0, 720, 720,
                              false,
                              720, 720,
                              false,
                              1, 720,
                              false };

    librtf::set_sectionformat(&sf);

    // Set default RTF paragraph formatting properties
    RTF_PARAGRAPH_FORMAT pf = { RTF_PARAGRAPHBREAK_NONE,
                                false, true,
                                RTF_PARAGRAPHALIGN_LEFT,
                                0, 0, 0, 0, 0, 0,
                                NULL,
                                false, false, false, false, false, false };

    pf.BORDERS.borderColor = 0;
    pf.BORDERS.borderKind = RTF_PARAGRAPHBORDERKIND_NONE;
    pf.BORDERS.borderSpace = 0;
    pf.BORDERS.borderType = RTF_PARAGRAPHBORDERTYPE_STHICK;
    pf.BORDERS.borderWidth = 0;

    pf.CHARACTER.animatedCharacter = false;
    pf.CHARACTER.backgroundColor = 0;
    pf.CHARACTER.boldCharacter = false;
    pf.CHARACTER.capitalCharacter = false;
    pf.CHARACTER.doublestrikeCharacter = false;
    pf.CHARACTER.embossCharacter = false;
    pf.CHARACTER.engraveCharacter = false;
    pf.CHARACTER.expandCharacter = 0;
    pf.CHARACTER.fontNumber = 0;
    pf.CHARACTER.fontSize = 24;
    pf.CHARACTER.foregroundColor = 0;
    pf.CHARACTER.italicCharacter = false;
    pf.CHARACTER.kerningCharacter = 0;
    pf.CHARACTER.outlineCharacter = false;
    pf.CHARACTER.scaleCharacter = 100;
    pf.CHARACTER.shadowCharacter = false;
    pf.CHARACTER.smallcapitalCharacter = false;
    pf.CHARACTER.strikeCharacter = false;
    pf.CHARACTER.subscriptCharacter = false;
    pf.CHARACTER.superscriptCharacter = false;
    pf.CHARACTER.underlineCharacter = 0;

    pf.NUMS.numsChar = char(0x95);
    pf.NUMS.numsLevel = 11;
    pf.NUMS.numsSpace = 360;

    pf.SHADING.shadingBkColor = 0;
    pf.SHADING.shadingFillColor = 0;
    pf.SHADING.shadingIntensity = 0;
    pf.SHADING.shadingType = RTF_PARAGRAPHSHADINGTYPE_FILL;

    pf.TABS.tabKind = RTF_PARAGRAPHTABKIND_NONE;
    pf.TABS.tabLead = RTF_PARAGRAPHTABLEAD_NONE;
    pf.TABS.tabPosition = 0;

    librtf::set_paragraphformat( &pf );

    // Set default RTF table row formatting properties
    RTF_TABLEROW_FORMAT rf = { RTF_ROWTEXTALIGN_LEFT,
                               0, 0, 0, 0, 0, 0 };

    librtf::set_tablerowformat( &rf );

    // Set default RTF table cell formatting properties
    RTF_TABLECELL_FORMAT cf = { RTF_CELLTEXTALIGN_CENTER,
                                0, 0, 0, 0,
                                RTF_CELLTEXTDIRECTION_LRTB,
                                false };

    cf.SHADING.shadingBkColor = 0;
    cf.SHADING.shadingFillColor = 0;
    cf.SHADING.shadingIntensity = 0;
    cf.SHADING.shadingType = RTF_PARAGRAPHSHADINGTYPE_FILL;

    cf.borderBottom.border = false;
    cf.borderBottom.BORDERS.borderColor = 0;
    cf.borderBottom.BORDERS.borderKind = RTF_PARAGRAPHBORDERKIND_NONE;
    cf.borderBottom.BORDERS.borderSpace = 0;
    cf.borderBottom.BORDERS.borderType = RTF_PARAGRAPHBORDERTYPE_STHICK;
    cf.borderBottom.BORDERS.borderWidth = 5;

    cf.borderLeft.border = false;
    cf.borderLeft.BORDERS.borderColor = 0;
    cf.borderLeft.BORDERS.borderKind = RTF_PARAGRAPHBORDERKIND_NONE;
    cf.borderLeft.BORDERS.borderSpace = 0;
    cf.borderLeft.BORDERS.borderType = RTF_PARAGRAPHBORDERTYPE_STHICK;
    cf.borderLeft.BORDERS.borderWidth = 5;

    cf.borderRight.border = false;
    cf.borderRight.BORDERS.borderColor = 0;
    cf.borderRight.BORDERS.borderKind = RTF_PARAGRAPHBORDERKIND_NONE;
    cf.borderRight.BORDERS.borderSpace = 0;
    cf.borderRight.BORDERS.borderType = RTF_PARAGRAPHBORDERTYPE_STHICK;
    cf.borderRight.BORDERS.borderWidth = 5;

    cf.borderTop.border = false;
    cf.borderTop.BORDERS.borderColor = 0;
    cf.borderTop.BORDERS.borderKind = RTF_PARAGRAPHBORDERKIND_NONE;
    cf.borderTop.BORDERS.borderSpace = 0;
    cf.borderTop.BORDERS.borderType = RTF_PARAGRAPHBORDERTYPE_STHICK;
    cf.borderTop.BORDERS.borderWidth = 5;

    librtf::set_tablecellformat( &cf );
}

// Sets new RTF document font table
void librtf::set_fonttable( const char* fonts )
{
    if ( fonts == NULL )
        return;

    char* tmpbsrc = strdup( fonts );

    if ( tmpbsrc == NULL )
        return;

    char* tmpb = tmpbsrc;

    // Clear old font table
    if ( rtfFontTable.size() > 0 )
        rtfFontTable.clear();

    // Set separator list
    char separator[] = ";";

    // Create new RTF document font table
    int   font_number = 0;
    char  font_table_entry[1024] = {0};
    char* token = strtok( tmpb, separator );

    while ( token != NULL )
    {
        // Format font table entry
        snprintf( font_table_entry, 1024,
                  "{\\f%d\\fnil\\fcharset0\\cpg1252 %s}",
                  font_number, token );

        rtfFontTable += font_table_entry;

        // Get next font
        token = strtok( NULL, separator );
        font_number++;
    }

    delete[] tmpbsrc;
}

// Sets new RTF document color table
void librtf::set_colortable( const char* colors )
{
    if ( colors == NULL )
        return;

    // Clear old color table
    if ( rtfColorTable.size() > 0 )
        rtfColorTable.clear();

    char* tmpbsrc = strdup( colors );

    if ( tmpbsrc == NULL )
        return;

    char* tmpb = tmpbsrc;

    // Set separator list
    char separator[] = ";";

    // Create new RTF document color table
    int   color_number = 0;
    char  color_table_entry[40] = {0};
    char* token = strtok( tmpb, separator );

    while ( token != NULL )
    {
        // Red
        snprintf( color_table_entry, 40, "\\red%s", token );
        rtfColorTable += color_table_entry;

        // Green
        token = strtok( NULL, separator );
        if ( token != NULL )
        {
            snprintf( color_table_entry, 40, "\\green%s", token );
            rtfColorTable += color_table_entry;
        }

        // Blue
        token = strtok( NULL, separator );
        if ( token != NULL )
        {
            snprintf( color_table_entry, 40, "\\blue%s;", token );
            rtfColorTable += color_table_entry;
        }

        // Get next color
        token = strtok( NULL, separator );
        color_number++;
    }

    delete[] tmpbsrc;
}

// Sets RTF document formatting properties
void librtf::set_documentformat( RTF_DOCUMENT_FORMAT* df )
{
    if ( df != NULL )
    {
        // Set new RTF document formatting properties
        memcpy( &rtfDocFormat, df, sizeof(RTF_DOCUMENT_FORMAT) );
    }
}

// Writes RTF document formatting properties
bool librtf::write_documentformat()
{
    // Set error flag
    bool result = true;

    // RTF document text
    char rtfText[1024] = {0};

    snprintf( rtfText, 1024,
              "\\viewkind%d\\viewscale%d\\paperw%d\\paperh%d\\margl%d\\margr%d\\margt%d\\margb%d\\gutter%d",
              rtfDocFormat.viewKind,
              rtfDocFormat.viewScale,
              rtfDocFormat.paperWidth,
              rtfDocFormat.paperHeight,
              rtfDocFormat.marginLeft,
              rtfDocFormat.marginRight,
              rtfDocFormat.marginTop,
              rtfDocFormat.marginBottom,
              rtfDocFormat.gutterWidth );

    if ( rtfDocFormat.facingPages )
        strcats( rtfText, "\\facingp", 1024 );

    if ( rtfDocFormat.readOnly )
        strcats( rtfText, "\\annotprot", 1024 );

    // Writes RTF document formatting properties
    if ( rtfFile != NULL )
    {
        if ( fwrite( rtfText, 1, strlen(rtfText), rtfFile ) < strlen(rtfText) )
            result = false;
    }
    else
    {
        result = false;
    }

    // Return error flag
    return result;
}

// Sets RTF section formatting properties
void librtf::set_sectionformat(RTF_SECTION_FORMAT* sf)
{
    if ( sf != NULL )
    {
        // Set new RTF section formatting properties
        memcpy( &rtfSecFormat, sf, sizeof(RTF_SECTION_FORMAT) );
    }
}

// Writes RTF section formatting properties
bool librtf::write_sectionformat()
{
    // Set error flag
    bool result = true;

    // RTF document text
    char rtfText[1024] = {0};

    // Format new section
    char text[1024] = {0};
    char pgn[100] = {0};

    if ( rtfSecFormat.newSection )
        strcats( text, "\\sect", 1024 );

    if ( rtfSecFormat.defaultSection )
        strcats( text, "\\sectd", 1024 );

    if ( rtfSecFormat.showPageNumber )
    {
        snprintf( pgn, 100, "\\pgnx%d\\pgny%d",
                  rtfSecFormat.pageNumberOffsetX,
                  rtfSecFormat.pageNumberOffsetY );

        strcats( text, pgn, 1024 );
    }

    // Format section break
    char sbr[100] = {0};

    switch ( rtfSecFormat.sectionBreak )
    {
        // Continuous break
        case RTF_SECTIONBREAK_CONTINUOUS:
            strcats( sbr, "\\sbknone", 100 );
            break;

        // Column break
        case RTF_SECTIONBREAK_COLUMN:
            strcats( sbr, "\\sbkcol", 100 );
            break;

        // Page break
        case RTF_SECTIONBREAK_PAGE:
            strcats( sbr, "\\sbkpage", 100 );
            break;

        // Even-page break
        case RTF_SECTIONBREAK_EVENPAGE:
            strcats( sbr, "\\sbkeven", 100 );
            break;

        // Odd-page break
        case RTF_SECTIONBREAK_ODDPAGE:
            strcats( sbr, "\\sbkodd", 100 );
            break;
    }

    // Format section columns
    char cols[100] = {0};
    if ( rtfSecFormat.cols == true )
    {
        // Format columns
        snprintf( cols, 100, "\\cols%d\\colsx%d",
                  rtfSecFormat.colsNumber, rtfSecFormat.colsDistance );

        if ( rtfSecFormat.colsLineBetween )
            strcats( cols, "\\linebetcol", 100 );
    }

    snprintf( rtfText, 1024,
              "\n%s%s%s\\pgwsxn%d\\pghsxn%d\\marglsxn%d\\margrsxn%d\\margtsxn%d\\margbsxn%d\\guttersxn%d\\headery%d\\footery%d",
              text, sbr, cols,
              rtfSecFormat.pageWidth, rtfSecFormat.pageHeight,
              rtfSecFormat.pageMarginLeft, rtfSecFormat.pageMarginRight,
              rtfSecFormat.pageMarginTop, rtfSecFormat.pageMarginBottom,
              rtfSecFormat.pageGutterWidth, rtfSecFormat.pageHeaderOffset,
              rtfSecFormat.pageFooterOffset );

    // Writes RTF section formatting properties
    if ( rtfFile != NULL )
    {
        if ( fwrite( rtfText, 1, strlen(rtfText), rtfFile ) < strlen(rtfText) )
            result = false;
    }
    else
    {
        result = false;
    }

    // Return error flag
    return result;
}


// Starts new RTF section
RTF_ERROR_TYPE librtf::start_section()
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_SUCCESS;

    // Set new section flag
    rtfSecFormat.newSection = true;

    // Starts new RTF section
    if( librtf::write_sectionformat() == false )
        error = RTF_SECTIONFORMAT_ERROR;

    // Return error flag
    return error;
}

// Sets RTF paragraph formatting properties
void librtf::set_paragraphformat(RTF_PARAGRAPH_FORMAT* pf)
{
    if ( pf != NULL )
    {
        // Set new RTF paragraph formatting properties
        memcpy( &rtfParFormat, pf, sizeof(RTF_PARAGRAPH_FORMAT) );
    }
}

// Writes RTF paragraph formatting properties
bool librtf::write_paragraphformat()
{
    // Set error flag
    bool result = true;

    // RTF document text
    char rtfText[4096] = {0};

    // Format new paragraph
    string text;

    if ( rtfParFormat.newParagraph )
        text += "\\par";

    if ( rtfParFormat.defaultParagraph )
        text += "\\pard";

    if ( rtfParFormat.tableText == false )
        text += "\\plain";
    else
        text += "\\intbl";

    switch ( rtfParFormat.paragraphBreak )
    {
        // No break
        case RTF_PARAGRAPHBREAK_NONE:
            break;

        // Page break;
        case RTF_PARAGRAPHBREAK_PAGE:
            text += "\\page";
            break;

        // Column break;
        case RTF_PARAGRAPHBREAK_COLUMN:
            text += "\\column";
            break;

        // Line break;
        case RTF_PARAGRAPHBREAK_LINE:
            text += "\\line";
            break;
    }

    // Format aligment
    switch ( rtfParFormat.paragraphAligment )
    {
        // Left aligned paragraph
        case RTF_PARAGRAPHALIGN_LEFT:
            text += "\\ql";
            break;

        // Center aligned paragraph
        case RTF_PARAGRAPHALIGN_CENTER:
            text += "\\qc";
            break;

        // Right aligned paragraph
        case RTF_PARAGRAPHALIGN_RIGHT:
            text += "\\qr";
            break;

        // Justified aligned paragraph
        case RTF_PARAGRAPHALIGN_JUSTIFY:
            text += "\\qj";
            break;
    }

    // Format tabs
    if ( rtfParFormat.paragraphTabs == true )
    {
        // Set tab kind
        switch ( rtfParFormat.TABS.tabKind )
        {
            // No tab
            case RTF_PARAGRAPHTABKIND_NONE:
                break;

            // Centered tab
            case RTF_PARAGRAPHTABKIND_CENTER:
                text += "\\tqc";
                break;

            // Flush-right tab
            case RTF_PARAGRAPHTABKIND_RIGHT:
                text += "\\tqr";
                break;

            // Decimal tab
            case RTF_PARAGRAPHTABKIND_DECIMAL:
                text += "\\tqdec";
                break;
        }

        // Set tab leader
        switch ( rtfParFormat.TABS.tabLead )
        {
            // No lead
            case RTF_PARAGRAPHTABLEAD_NONE:
                break;

            // Leader dots
            case RTF_PARAGRAPHTABLEAD_DOT:
                text += "\\tldot";
                break;

            // Leader middle dots
            case RTF_PARAGRAPHTABLEAD_MDOT:
                text += "\\tlmdot";
                break;

            // Leader hyphens
            case RTF_PARAGRAPHTABLEAD_HYPH:
                text += "\\tlhyph";
                break;

            // Leader underline
            case RTF_PARAGRAPHTABLEAD_UNDERLINE:
                text += "\\tlul";
                break;

            // Leader thick line
            case RTF_PARAGRAPHTABLEAD_THICKLINE:
                text += "\\tlth";
                break;

            // Leader equal sign
            case RTF_PARAGRAPHTABLEAD_EQUAL:
                text += "\\tleq";
                break;
        }

        // Set tab position
        char tb[20] = {0};
        snprintf( tb, 20, "\\tx%d", rtfParFormat.TABS.tabPosition );
        text += tb;
    }

    // Format bullets and numbering
    if ( rtfParFormat.paragraphNums == true )
    {
        char nums[80] = {0};
        snprintf( nums, 80,
                  "{\\*\\pn\\pnlvl%d\\pnsp%d\\pntxtb %c}",
                  rtfParFormat.NUMS.numsLevel,
                  rtfParFormat.NUMS.numsSpace,
                  rtfParFormat.NUMS.numsChar );
        text += nums;
    }

    // Format paragraph borders
    if ( rtfParFormat.paragraphBorders == true )
    {
        string border;

        // Format paragraph border kind
        switch (rtfParFormat.BORDERS.borderKind)
        {
            // No border
            case RTF_PARAGRAPHBORDERKIND_NONE:
                break;

            // Border top
            case RTF_PARAGRAPHBORDERKIND_TOP:
                border += "\\brdrt";
                break;

            // Border bottom
            case RTF_PARAGRAPHBORDERKIND_BOTTOM:
                border += "\\brdrb";
                break;

            // Border left
            case RTF_PARAGRAPHBORDERKIND_LEFT:
                border += "\\brdrl";
                break;

            // Border right
            case RTF_PARAGRAPHBORDERKIND_RIGHT:
                border += "\\brdrr";
                break;

            // Border box
            case RTF_PARAGRAPHBORDERKIND_BOX:
                border += "\\box";
                break;
        }

        // Format paragraph border type
        const char *br = librtf::get_bordername( rtfParFormat.BORDERS.borderType );
        if ( br != NULL )
        {
            border += br;
        }

        // Set paragraph border width
        char brd[100] = {0};
        snprintf( brd, 100, "\\brdrw%d\\brsp%d",
                  rtfParFormat.BORDERS.borderWidth,
                  rtfParFormat.BORDERS.borderSpace );
        border += brd;
        text += border;

        // Set paragraph border color
        char brdcol[100] = {0};
        snprintf( brdcol, 100, "\\brdrcf%d",
                  rtfParFormat.BORDERS.borderColor );
        text += brdcol;
    }

    // Format paragraph shading
    if ( rtfParFormat.paragraphShading == true )
    {
        char shading[40] = {0};
        snprintf( shading, 40, "\\shading%d", rtfParFormat.SHADING.shadingIntensity );

        // Format paragraph shading
        const char* sh = librtf::get_shadingname( rtfParFormat.SHADING.shadingType, false );
        if ( sh != NULL )
            text += sh;

        // Set paragraph shading color
        char shcol[100] = {0};
        snprintf( shcol, 100,
                  "\\cfpat%d\\cbpat%d",
                  rtfParFormat.SHADING.shadingFillColor,
                  rtfParFormat.SHADING.shadingBkColor );
        text += shcol;
    }

    // Format paragraph font
    string font;
    char tmps[1024] = {0};

    snprintf( tmps, 1024,
              "\\animtext%d\\expndtw%d\\kerning%d\\charscalex%d\\f%d\\fs%d\\cf%d",
              rtfParFormat.CHARACTER.animatedCharacter,
              rtfParFormat.CHARACTER.expandCharacter,
              rtfParFormat.CHARACTER.kerningCharacter,
              rtfParFormat.CHARACTER.scaleCharacter,
              rtfParFormat.CHARACTER.fontNumber,
              rtfParFormat.CHARACTER.fontSize,
              rtfParFormat.CHARACTER.foregroundColor );

    font += tmps;

    if ( rtfParFormat.CHARACTER.boldCharacter )
        font += "\\b";
    else
        font += "\\b0";

    if ( rtfParFormat.CHARACTER.capitalCharacter )
        font += "\\caps";
    else
        font += "\\caps0";

    if ( rtfParFormat.CHARACTER.doublestrikeCharacter )
        font += "\\striked1";
    else
        font += "\\striked0";

    if ( rtfParFormat.CHARACTER.embossCharacter )
        font += "\\embo";
    if ( rtfParFormat.CHARACTER.engraveCharacter )
        font += "\\impr";

    if ( rtfParFormat.CHARACTER.italicCharacter )
        font += "\\i";
    else
        font += "\\i0";

    if ( rtfParFormat.CHARACTER.outlineCharacter )
        font += "\\outl";
    else
        font += "\\outl0";

    if ( rtfParFormat.CHARACTER.shadowCharacter )
        font += "\\shad";
    else
        font += "\\shad0";

    if ( rtfParFormat.CHARACTER.smallcapitalCharacter )
        font += "\\scaps";
    else
        font += "\\scaps0";

    if ( rtfParFormat.CHARACTER.strikeCharacter )
        font += "\\strike";
    else
        font += "\\strike0";

    if ( rtfParFormat.CHARACTER.subscriptCharacter )
        font += "\\sub";

    if ( rtfParFormat.CHARACTER.superscriptCharacter )
        font += "\\super";

    switch (rtfParFormat.CHARACTER.underlineCharacter)
    {
        // None underline
        case 0:
            font += "\\ulnone";
            break;

        // Continuous underline
        case 1:
            font += "\\ul";
            break;

        // Dotted underline
        case 2:
            font += "\\uld";
            break;

        // Dashed underline
        case 3:
            font += "\\uldash";
            break;

        // Dash-dotted underline
        case 4:
            font += "\\uldashd";
            break;

        // Dash-dot-dotted underline
        case 5:
            font += "\\uldashdd";
            break;

        // Double underline
        case 6:
            font += "\\uldb";
            break;

        // Heavy wave underline
        case 7:
            font += "\\ulhwave";
            break;

        // Long dashed underline
        case 8:
            font += "\\ulldash";
            break;

        // Thick underline
        case 9:
            font += "\\ulth";
            break;

        // Thick dotted underline
        case 10:
            font += "\\ulthd";
            break;

        // Thick dashed underline
        case 11:
            font += "\\ulthdash";
            break;

        // Thick dash-dotted underline
        case 12:
            font += "\\ulthdashd";
            break;

        // Thick dash-dot-dotted underline
        case 13:
            font += "\\ulthdashdd";
            break;

        // Thick long dashed underline
        case 14:
            font += "\\ulthldash";
            break;

        // Double wave underline
        case 15:
            font += "\\ululdbwave";
            break;

        // Word underline
        case 16:
            font += "\\ulw";
            break;

        // Wave underline
        case 17:
            font += "\\ulwave";
            break;
    }

    // Set paragraph tabbed text
    if ( rtfParFormat.paragraphText != NULL )
    {
        if ( rtfParFormat.tabbedText == false )
        {
            snprintf( rtfText, 4096,
                      "\n%s\\fi%d\\li%d\\ri%d\\sb%d\\sa%d\\sl%d%s %s",
                      text.c_str(),
                      rtfParFormat.firstLineIndent,
                      rtfParFormat.leftIndent,
                      rtfParFormat.rightIndent,
                      rtfParFormat.spaceBefore,
                      rtfParFormat.spaceAfter,
                      rtfParFormat.lineSpacing,
                      font.c_str(),
                      rtfParFormat.paragraphText );
        }
        else
        {
            snprintf( rtfText, 4096, "\\tab %s",
                      rtfParFormat.paragraphText );
        }
    }

    // Writes RTF paragraph formatting properties
    if ( rtfFile != NULL )
    {
        if ( fwrite( rtfText, 1, strlen(rtfText), rtfFile ) < strlen(rtfText) )
            result = false;
    }
    else
    {
        result = false;
    }

    // Return error flag
    return result;
}

// Starts new RTF paragraph
RTF_ERROR_TYPE librtf::start_paragraph( const char* text, bool newPar )
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_ERROR;

    if ( rtfParFormat.paragraphText != NULL )
    {
        delete[] rtfParFormat.paragraphText;
        rtfParFormat.paragraphText = NULL;
    }

    if ( text != NULL )
    {
        // Copy paragraph text
        size_t sl = strlen(text);
        rtfParFormat.paragraphText = new char[sl+1];

        if ( rtfParFormat.paragraphText != NULL )
        {
            memset( rtfParFormat.paragraphText, 0, sl+1 );
            memcpy( rtfParFormat.paragraphText, text, sl );

            // Set new paragraph
            rtfParFormat.newParagraph = newPar;

            // Starts new RTF paragraph
            if( librtf::write_paragraphformat() == false )
                error = RTF_PARAGRAPHFORMAT_ERROR;
            else
                error = RTF_SUCCESS;
        }
    }

    // Return error flag
    return error;
}

// Gets RTF document formatting properties
RTF_DOCUMENT_FORMAT* librtf::get_documentformat()
{
    // Get current RTF document formatting properties
    return &rtfDocFormat;
}

// Gets RTF section formatting properties
RTF_SECTION_FORMAT* librtf::get_sectionformat()
{
    // Get current RTF section formatting properties
    return &rtfSecFormat;
}

// Gets RTF paragraph formatting properties
RTF_PARAGRAPH_FORMAT* librtf::get_paragraphformat()
{
    // Get current RTF paragraph formatting properties
    return &rtfParFormat;
}

// Loads image from file
RTF_ERROR_TYPE librtf::load_image( const char* image, int width, int height )
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_FAILURE;

    if ( rtfFile == NULL )
        return error;

    // Check image type by file extension
    bool err = false;

    // If valid image type
    // Free IPicture object
    if ( rtfPicture != NULL )
    {
        rtfPicture->Release();
        rtfPicture = NULL;
    }

    // Read image file
    int imageFile = _open( image, _O_RDONLY | _O_BINARY );
    struct _stat st;
    _fstat( imageFile, &st );
    DWORD nSize = st.st_size;
    BYTE* pBuff = new BYTE[nSize];
    _read( imageFile, pBuff, nSize );

    // Alocate memory for image data
    HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, nSize);
    void* pData = GlobalLock(hGlobal);
    memcpy(pData, pBuff, nSize);
    GlobalUnlock(hGlobal);

    // Load image using OLE
    IStream* pStream = NULL;

    if ( CreateStreamOnHGlobal(hGlobal, TRUE, &pStream) == S_OK )
    {
        HRESULT hr;

        if ((hr = OleLoadPicture( pStream, nSize, FALSE, IID_IPicture, (LPVOID *)&rtfPicture)) != S_OK)
            error = RTF_IMAGE_ERROR;

        pStream->Release();
    }

    delete []pBuff;
    _close(imageFile);

    // If image is loaded
    if ( rtfPicture != NULL )
    {
        // Calculate image size
        long hmWidth = 0;
        long hmHeight = 0;
        rtfPicture->get_Width(&hmWidth);
        rtfPicture->get_Height(&hmHeight);
        int nWidth  = MulDiv( hmWidth, GetDeviceCaps(GetDC(NULL),LOGPIXELSX), 2540 );
        int nHeight = MulDiv( hmHeight, GetDeviceCaps(GetDC(NULL),LOGPIXELSY), 2540 );

        // Create metafile;
        HDC hdcMeta = CreateMetaFile(NULL);

        // Render picture to metafile
        rtfPicture->Render( hdcMeta, 0, 0, nWidth, nHeight, 0, hmHeight, hmWidth, -hmHeight, NULL );

        // Close metafile
        HMETAFILE hmf = CloseMetaFile(hdcMeta);

        // Get metafile data
        UINT size = GetMetaFileBitsEx( hmf, 0, NULL );
        BYTE* buffer = new BYTE[size];
        GetMetaFileBitsEx( hmf, size, buffer );
        DeleteMetaFile(hmf);

        // Convert metafile binary data to hexadecimal
        char* hexstr = librtf::bin_hex_convert( buffer, size );
        delete []buffer;

        // Format picture paragraph
        RTF_PARAGRAPH_FORMAT* pf = librtf::get_paragraphformat();
        pf->paragraphText = NULL;
        librtf::write_paragraphformat();

        // Writes RTF picture data
        char rtfText[128] = {0};
        snprintf( rtfText, 128,
                  "\n{\\pict\\wmetafile8\\picwgoal%d\\pichgoal%d\\picscalex%d\\picscaley%d\n",
                  hmWidth, hmHeight, width, height );

        if ( fwrite( rtfText, 1, strlen(rtfText), rtfFile ) < strlen(rtfText) )
        {
            error = RTF_IMAGE_ERROR;
            return error;
        }

        fwrite( hexstr, 1, 2*size, rtfFile );

        delete[] hexstr;

        strncpy( rtfText, "}", 128 );
        fwrite( rtfText, 1, strlen(rtfText), rtfFile );

        error = RTF_SUCCESS;
    }

    // Return error flag
    return error;
}

// Converts binary data to hex
char* librtf::bin_hex_convert( const unsigned char* binary, size_t size )
{
    if ( (binary == NULL ) || ( size == 0 ) )
        return NULL;

    size_t actualsize = (2 * size) + 1;
    char* result = new char[ actualsize ];

    if ( result == NULL )
        return NULL;

    memset( result, 0, actualsize );

    char part1, part2;

    for ( size_t cnt=0; cnt<size; cnt++ )
    {
        part1 = binary[cnt] / 16;

        if ( part1 < 10 )
            part1 += 48;
        else
        {
            if ( part1 == 10 )
                part1 = 'a';
            if ( part1 == 11 )
                part1 = 'b';
            if ( part1 == 12 )
                part1 = 'c';
            if ( part1 == 13 )
                part1 = 'd';
            if ( part1 == 14 )
                part1 = 'e';
            if ( part1 == 15 )
                part1 = 'f';
        }

        part2 = binary[cnt] % 16;

        if ( part2 < 10 )
            part2 += 48;
        else
        {
            if ( part2 == 10 )
                part2 = 'a';
            if ( part2 == 11 )
                part2 = 'b';
            if ( part2 == 12 )
                part2 = 'c';
            if ( part2 == 13 )
                part2 = 'd';
            if ( part2 == 14 )
                part2 = 'e';
            if ( part2 == 15 )
                part2 = 'f';
        }

        result[2*cnt  ] = part1;
        result[2*cnt+1] = part2;
    }

    return result;
}


// Starts new RTF table row
RTF_ERROR_TYPE librtf::start_tablerow()
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_SUCCESS;

    string tblrw;

    // Format table row aligment
    switch (rtfRowFormat.rowAligment)
    {
        // Left align
        case RTF_ROWTEXTALIGN_LEFT:
            tblrw = "\\trql";
            break;

        // Center align
        case RTF_ROWTEXTALIGN_CENTER:
            tblrw = "\\trqc";
            break;

        // Right align
        case RTF_ROWTEXTALIGN_RIGHT:
            tblrw = "\\trqr";
            break;
    }

    // Writes RTF table data
    char rtfText[1024] = {0};

    snprintf( rtfText, 1024,
              "\n\\trowd\\trgaph115%s\\trleft%d\\trrh%d\\trpaddb%d\\trpaddfb3\\trpaddl%d\\trpaddfl3\\trpaddr%d\\trpaddfr3\\trpaddt%d\\trpaddft3",
              tblrw.c_str(),
              rtfRowFormat.rowLeftMargin,
              rtfRowFormat.rowHeight,
              rtfRowFormat.marginTop,
              rtfRowFormat.marginBottom,
              rtfRowFormat.marginLeft,
              rtfRowFormat.marginRight );

    if ( rtfFile != NULL )
    {
        if ( fwrite( rtfText, 1, strlen(rtfText), rtfFile ) < strlen(rtfText) )
            error = RTF_TABLE_ERROR;
    }
    else
    {
        error = RTF_FAILURE;
    }

    // Return error flag
    return error;
}


// Ends RTF table row
RTF_ERROR_TYPE librtf::end_tablerow()
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_SUCCESS;

    // Writes RTF table data
    char rtfText[] = "\n\\trgaph115\\row\\pard";

    if ( rtfFile != NULL )
    {
        if ( fwrite( rtfText, 1, strlen(rtfText), rtfFile ) < strlen(rtfText) )
            error = RTF_TABLE_ERROR;
    }
    else
    {
        error = RTF_FAILURE;
    }

    // Return error flag
    return error;
}


// Starts new RTF table cell
RTF_ERROR_TYPE librtf::start_tablecell(int rightMargin)
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_SUCCESS;

    char tblcla[20] = {0};

    // Format table cell text aligment
    switch (rtfCellFormat.textVerticalAligment)
    {
        // Top align
        case RTF_CELLTEXTALIGN_TOP:
            strncpy( tblcla, "\\clvertalt", 20 );
            break;

        // Center align
        case RTF_CELLTEXTALIGN_CENTER:
            strncpy( tblcla, "\\clvertalc", 20 );
            break;

        // Bottom align
        case RTF_CELLTEXTALIGN_BOTTOM:
            strncpy( tblcla, "\\clvertalb", 20 );
            break;
    }

    char tblcld[20] = {0};

    // Format table cell text direction
    switch (rtfCellFormat.textDirection)
    {
        // Left to right, top to bottom
        case RTF_CELLTEXTDIRECTION_LRTB:
            strncpy( tblcld, "\\cltxlrtb", 20 );
            break;

        // Right to left, top to bottom
        case RTF_CELLTEXTDIRECTION_RLTB:
            strncpy( tblcld, "\\cltxtbrl", 20 );
            break;

        // Left to right, bottom to top
        case RTF_CELLTEXTDIRECTION_LRBT:
            strncpy( tblcld, "\\cltxbtlr", 20 );
            break;

        // Left to right, top to bottom, vertical
        case RTF_CELLTEXTDIRECTION_LRTBV:
            strncpy( tblcld, "\\cltxlrtbv", 20 );
            break;

        // Right to left, top to bottom, vertical
        case RTF_CELLTEXTDIRECTION_RLTBV:
            strncpy( tblcld, "\\cltxtbrlv", 20 );
            break;
    }

    char tbclbrb[1024] = {0};
    char tbclbrl[1024] = {0};
    char tbclbrr[1024] = {0};
    char tbclbrt[1024] = {0};

    // Format table cell border
    if ( rtfCellFormat.borderBottom.border == true )
    {
        // Bottom cell border
        const char* border = librtf::get_bordername(rtfCellFormat.borderBottom.BORDERS.borderType);

        if ( border != NULL )
        {
            snprintf( tbclbrb, 1024,
                      "%s%s\\brdrw%d\\brsp%d\\brdrcf%d",
                      "\\clbrdrb",
                      border,
                      rtfCellFormat.borderBottom.BORDERS.borderWidth,
                      rtfCellFormat.borderBottom.BORDERS.borderSpace,
                      rtfCellFormat.borderBottom.BORDERS.borderColor );
        }
    }

    if ( rtfCellFormat.borderLeft.border == true )
    {
        // Left cell border
        const char* border = librtf::get_bordername(rtfCellFormat.borderLeft.BORDERS.borderType);

        if ( border != NULL )
        {
            snprintf( tbclbrl, 1024,
                      "%s%s\\brdrw%d\\brsp%d\\brdrcf%d",
                      "\\clbrdrl",
                      border,
                      rtfCellFormat.borderLeft.BORDERS.borderWidth,
                      rtfCellFormat.borderLeft.BORDERS.borderSpace,
                      rtfCellFormat.borderLeft.BORDERS.borderColor );
        }
    }
    if ( rtfCellFormat.borderRight.border == true )
    {
        // Right cell border
        const char* border = librtf::get_bordername(rtfCellFormat.borderRight.BORDERS.borderType);

        if ( border != NULL )
        {
            snprintf( tbclbrr, 1024,
                      "%s%s\\brdrw%d\\brsp%d\\brdrcf%d",
                      "\\clbrdrr",
                      border,
                      rtfCellFormat.borderRight.BORDERS.borderWidth,
                      rtfCellFormat.borderRight.BORDERS.borderSpace,
                      rtfCellFormat.borderRight.BORDERS.borderColor );
        }
    }
    if ( rtfCellFormat.borderTop.border == true )
    {
        // Top cell border
        const char* border = librtf::get_bordername(rtfCellFormat.borderTop.BORDERS.borderType);
        if ( border != NULL )
        {
            snprintf( tbclbrt, 1024,
                      "%s%s\\brdrw%d\\brsp%d\\brdrcf%d",
                      "\\clbrdrt",
                      border,
                      rtfCellFormat.borderTop.BORDERS.borderWidth,
                      rtfCellFormat.borderTop.BORDERS.borderSpace,
                      rtfCellFormat.borderTop.BORDERS.borderColor );
        }
    }

    // Format table cell shading
    char shading[128] = {0};

    if ( rtfCellFormat.cellShading == true )
    {
        const char* sh = librtf::get_shadingname( rtfCellFormat.SHADING.shadingType, true );

        if ( sh != NULL )
        {
            // Set paragraph shading color
            snprintf( shading, 128,
                      "%s\\clshdgn%d\\clcfpat%d\\clcbpat%d",
                      sh,
                      rtfCellFormat.SHADING.shadingIntensity,
                      rtfCellFormat.SHADING.shadingFillColor,
                      rtfCellFormat.SHADING.shadingBkColor );
        }
    }

    // Writes RTF table data
    char rtfText[1024] = {0};
    snprintf( rtfText, 1024,
              "\n\\tcelld%s%s%s%s%s%s%s\\cellx%d",
              tblcla, tblcld, tbclbrb, tbclbrl, tbclbrr, tbclbrt,
              shading, rightMargin );

    if ( fwrite( rtfText, 1, strlen(rtfText), rtfFile ) < strlen(rtfText) )
        error = RTF_TABLE_ERROR;

    // Return error flag
    return error;
}

// Ends RTF table cell
RTF_ERROR_TYPE librtf::end_tablecell()
{
    // Set error flag
    RTF_ERROR_TYPE error = RTF_SUCCESS;

    // Writes RTF table data
    char rtfText[] = "\n\\cell ";

    if ( rtfFile != NULL )
    {
        if ( fwrite( rtfText, 1, strlen(rtfText), rtfFile ) <= 0 )
            error = RTF_TABLE_ERROR;
    }
    else
    {
        error = RTF_FAILURE;
    }

    // Return error flag
    return error;
}


// Gets RTF table row formatting properties
RTF_TABLEROW_FORMAT* librtf::get_tablerowformat()
{
    // Get current RTF table row formatting properties
    return &rtfRowFormat;
}


// Sets RTF table row formatting properties
void librtf::set_tablerowformat(RTF_TABLEROW_FORMAT* rf)
{
    // Set new RTF table row formatting properties
    if ( rf != NULL )
        memcpy( &rtfRowFormat, rf, sizeof(RTF_TABLEROW_FORMAT) );
}


// Gets RTF table cell formatting properties
RTF_TABLECELL_FORMAT* librtf::get_tablecellformat()
{
    // Get current RTF table cell formatting properties
    return &rtfCellFormat;
}


// Sets RTF table cell formatting properties
void librtf::set_tablecellformat(RTF_TABLECELL_FORMAT* cf)
{
    // Set new RTF table cell formatting properties
    if ( cf != NULL )
        memcpy( &rtfCellFormat, cf, sizeof(RTF_TABLECELL_FORMAT) );
}

// Gets border name
const char* librtf::get_bordername(int border_type)
{
    static string border;

    if ( border.size() > 0 )
        border.clear();

    switch (border_type)
    {
        // Single-thickness border
        case RTF_PARAGRAPHBORDERTYPE_STHICK:
            border = "\\brdrs";
            break;

        // Double-thickness border
        case RTF_PARAGRAPHBORDERTYPE_DTHICK:
            border = "\\brdrth";
            break;

        // Shadowed border
        case RTF_PARAGRAPHBORDERTYPE_SHADOW:
            border = "\\brdrsh";
            break;

        // Double border
        case RTF_PARAGRAPHBORDERTYPE_DOUBLE:
            border = "\\brdrdb";
            break;

        // Dotted border
        case RTF_PARAGRAPHBORDERTYPE_DOT:
            border = "\\brdrdot";
            break;

        // Dashed border
        case RTF_PARAGRAPHBORDERTYPE_DASH:
            border =  "\\brdrdash";
            break;

        // Hairline border
        case RTF_PARAGRAPHBORDERTYPE_HAIRLINE:
            border = "\\brdrhair";
            break;

        // Inset border
        case RTF_PARAGRAPHBORDERTYPE_INSET:
            border = "\\brdrinset";
            break;

        // Dashed border (small)
        case RTF_PARAGRAPHBORDERTYPE_SDASH:
            border = "\\brdrdashsm";
            break;

        // Dot-dashed border
        case RTF_PARAGRAPHBORDERTYPE_DOTDASH:
            border = "\\brdrdashd";
            break;

        // Dot-dot-dashed border
        case RTF_PARAGRAPHBORDERTYPE_DOTDOTDASH:
            border = "\\brdrdashdd";
            break;

        // Outset border
        case RTF_PARAGRAPHBORDERTYPE_OUTSET:
            border = "\\brdroutset";
            break;

        // Triple border
        case RTF_PARAGRAPHBORDERTYPE_TRIPLE:
            border = "\\brdrtriple";
            break;

        // Wavy border
        case RTF_PARAGRAPHBORDERTYPE_WAVY:
            border = "\\brdrwavy";
            break;

        // Double wavy border
        case RTF_PARAGRAPHBORDERTYPE_DWAVY:
            border = "\\brdrwavydb";
            break;

        // Striped border
        case RTF_PARAGRAPHBORDERTYPE_STRIPED:
            border = "\\brdrdashdotstr";
            break;

        // Embossed border
        case RTF_PARAGRAPHBORDERTYPE_EMBOSS:
            border = "\\brdremboss";
            break;

        // Engraved border
        case RTF_PARAGRAPHBORDERTYPE_ENGRAVE:
            border = "\\brdrengrave";
            break;
    }

    return border.c_str();
}


// Gets shading name
const char* librtf::get_shadingname( int shading_type, bool cell )
{
    static string shading;

    if( shading.size() > 0 )
        shading.clear();

    if ( cell == false )
    {
        switch (shading_type)
        {
            // Fill shading
            case RTF_PARAGRAPHSHADINGTYPE_FILL:
                break;

            // Horizontal background pattern
            case RTF_PARAGRAPHSHADINGTYPE_HORIZ:
                shading = "\\bghoriz";
                break;

            // Vertical background pattern
            case RTF_PARAGRAPHSHADINGTYPE_VERT:
                shading = "\\bgvert";
                break;

            // Forward diagonal background pattern
            case RTF_PARAGRAPHSHADINGTYPE_FDIAG:
                shading = "\\bgfdiag";
                break;

            // Backward diagonal background pattern
            case RTF_PARAGRAPHSHADINGTYPE_BDIAG:
                shading = "\\bgbdiag";
                break;

            // Cross background pattern
            case RTF_PARAGRAPHSHADINGTYPE_CROSS:
                shading = "\\bgcross";
                break;

            // Diagonal cross background pattern
            case RTF_PARAGRAPHSHADINGTYPE_CROSSD:
                shading = "\\bgdcross";
                break;

            // Dark horizontal background pattern
            case RTF_PARAGRAPHSHADINGTYPE_DHORIZ:
                shading = "\\bgdkhoriz";
                break;

            // Dark vertical background pattern
            case RTF_PARAGRAPHSHADINGTYPE_DVERT:
                shading = "\\bgdkvert";
                break;

            // Dark forward diagonal background pattern
            case RTF_PARAGRAPHSHADINGTYPE_DFDIAG:
                shading = "\\bgdkfdiag";
                break;

            // Dark backward diagonal background pattern
            case RTF_PARAGRAPHSHADINGTYPE_DBDIAG:
                shading = "\\bgdkbdiag";
                break;

            // Dark cross background pattern
            case RTF_PARAGRAPHSHADINGTYPE_DCROSS:
                shading = "\\bgdkcross";
                break;

            // Dark diagonal cross background pattern
            case RTF_PARAGRAPHSHADINGTYPE_DCROSSD:
                shading = "\\bgdkdcross";
                break;
        }
    }
    else
    {
        switch (shading_type)
        {
            // Fill shading
            case RTF_CELLSHADINGTYPE_FILL:
                break;

            // Horizontal background pattern
            case RTF_CELLSHADINGTYPE_HORIZ:
                shading = "\\clbghoriz";
                break;

            // Vertical background pattern
            case RTF_CELLSHADINGTYPE_VERT:
                shading = "\\clbgvert";
                break;

            // Forward diagonal background pattern
            case RTF_CELLSHADINGTYPE_FDIAG:
                shading = "\\clbgfdiag";
                break;

            // Backward diagonal background pattern
            case RTF_CELLSHADINGTYPE_BDIAG:
                shading = "\\clbgbdiag";
                break;

            // Cross background pattern
            case RTF_CELLSHADINGTYPE_CROSS:
                shading = "\\clbgcross";
                break;

            // Diagonal cross background pattern
            case RTF_CELLSHADINGTYPE_CROSSD:
                shading = "\\clbgdcross";
                break;

            // Dark horizontal background pattern
            case RTF_CELLSHADINGTYPE_DHORIZ:
                shading = "\\clbgdkhoriz";
                break;

            // Dark vertical background pattern
            case RTF_CELLSHADINGTYPE_DVERT:
                shading = "\\clbgdkvert";
                break;

            // Dark forward diagonal background pattern
            case RTF_CELLSHADINGTYPE_DFDIAG:
                shading = "\\clbgdkfdiag";
                break;

            // Dark backward diagonal background pattern
            case RTF_CELLSHADINGTYPE_DBDIAG:
                shading = "\\clbgdkbdiag";
                break;

            // Dark cross background pattern
            case RTF_CELLSHADINGTYPE_DCROSS:
                shading = "\\clbgdkcross";
                break;

            // Dark diagonal cross background pattern
            case RTF_CELLSHADINGTYPE_DCROSSD:
                shading = "\\clbgdkdcross";
                break;
        }
    }

    return shading.c_str();
}
