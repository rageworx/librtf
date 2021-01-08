#ifndef __LIBRTF_H__
#define __LIBRTF_H__

#include "librtferrors.h"
#include "librtfdefines.h"
#include "librtfstructures.h"

// =============================================================================
// Original source at
//     https://www.codeproject.com/Articles/10582/rtflib-v1-0
// Original author : darkoman
//     https://www.codeproject.com/script/Membership/View.aspx?mid=183888
//
//  - Nikolić Darko, application developer
//  - Serbia & Montenegro
//  - 18000 Niš
//  - elmeh@bankerinter.net
//
// -----------------------------------------------------------------------------
// librtf v1.2.3 at https://github.com/rageworx/librtf
// An open source RTF wrting library for MinGW-W64.
// Sources rewritten by Raphael Kim
// =============================================================================

namespace librtf
{
    // Creates new RTF document
    RTF_ERROR_TYPE open( const char* filename = NULL,
                         const char* fonts = NULL,
                         const char* colors = NULL,
                         RTF_DOCUMENT_FORMAT* fmt = NULL );

    // Closes created RTF document
    RTF_ERROR_TYPE close();

    // Writes RTF document header
    bool write_header();

    // Sets global RTF library params
    void init();

    // Sets new RTF document font table
    void set_fonttable( const char* fonts );

    // Sets new RTF document color table
    void set_colortable( const char* colors );

    // Gets RTF document formatting properties
    RTF_DOCUMENT_FORMAT* get_documentformat();

    // Sets RTF document formatting properties
    void set_documentformat( RTF_DOCUMENT_FORMAT* df );

    // Writes RTF document formatting properties
    bool write_documentformat();

    // Gets RTF section formatting properties
    RTF_SECTION_FORMAT* get_sectionformat();

    // Sets RTF section formatting properties
    void set_sectionformat(RTF_SECTION_FORMAT* sf);

	// Writes RTF section formatting properties
    bool write_sectionformat();

    // Starts new RTF section
    RTF_ERROR_TYPE start_section();

    // Gets RTF paragraph formatting properties
    RTF_PARAGRAPH_FORMAT* get_paragraphformat();

    // Sets RTF paragraph formatting properties
    void set_paragraphformat( RTF_PARAGRAPH_FORMAT* pf );

    // Writes RTF paragraph formatting properties
    bool write_paragraphformat();

    // Starts new RTF paragraph
    RTF_ERROR_TYPE start_paragraph( const char* text, bool newPar );

    // Loads image from file
    RTF_ERROR_TYPE load_image( const char* image, int width, int height);

    // Converts binary data to hex
    char* bin_hex_convert( const unsigned char* binary, size_t size );

    // Sets default RTF document formatting
    void set_defaultformat();

    // Starts new RTF table row
    RTF_ERROR_TYPE  start_tablerow();

    // Ends RTF table row
    RTF_ERROR_TYPE  end_tablerow();

    // Starts new RTF table cell
    RTF_ERROR_TYPE  start_tablecell(int rightMargin);

    // Ends RTF table cell
    RTF_ERROR_TYPE  end_tablecell();

    // Gets RTF table row formatting properties
    RTF_TABLEROW_FORMAT* get_tablerowformat();

    // Sets RTF table row formatting properties
    void set_tablerowformat( RTF_TABLEROW_FORMAT* rf );

    // Gets RTF table cell formatting properties
    RTF_TABLECELL_FORMAT* get_tablecellformat();

    // Sets RTF table cell formatting properties
    void  set_tablecellformat( RTF_TABLECELL_FORMAT* cf );

    // Gets border name
    const char* get_bordername( int border_type );

    // Gets shading name
    const char* get_shadingname( int shading_type, bool cell );
};

#endif /// of __LIBRTF_H__
