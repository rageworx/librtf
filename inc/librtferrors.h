#ifndef __LIBRTF_ERRORS_H__
#define __LIBRTF_ERRORS_H__

#define RTF_ERROR_TYPE              unsigned short
#define RTF_ERROR                   0xFFFF
#define RTF_FAILURE                 RTF_ERROR
#define RTF_OPEN_ERROR				0x0001  /// Could not create RTF file
#define RTF_CLOSE_ERROR				0x0002	/// Could not close RTF file
#define RTF_HEADER_ERROR			0x0003	/// Could not write header part to RTF file
#define RTF_DOCUMENTFORMAT_ERROR	0x0004	/// Could not write document formatting properties to RTF file
#define RTF_SECTIONFORMAT_ERROR		0x0005	/// Could not write section formatting properties to RTF file
#define RTF_PARAGRAPHFORMAT_ERROR	0x0006	/// Could not write paragraph formatting properties to RTF file
#define RTF_IMAGE_ERROR				0x0007	/// Could not write image to RTF file
#define RTF_TABLE_ERROR				0x0008	/// Could not write table to RTF file
#define RTF_SUCCESS					0x1000	/// No error

#endif /// of __LIBRTF_ERRORS_H__