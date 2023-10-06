<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2013 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel_Shared
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Shared_Font
 *
 * @category   PHPExcel
 * @package    PHPExcel_Shared
 * @copyright  Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Shared_Font
{
	/* Methods for resolving autosize value */
	final public const AUTOSIZE_METHOD_APPROX	= 'approx';
	final public const AUTOSIZE_METHOD_EXACT		= 'exact';

	private static array $_autoSizeMethods = [self::AUTOSIZE_METHOD_APPROX, self::AUTOSIZE_METHOD_EXACT];

	/** Character set codes used by BIFF5-8 in Font records */
	final public const CHARSET_ANSI_LATIN				= 0x00;
	final public const CHARSET_SYSTEM_DEFAULT			= 0x01;
	final public const CHARSET_SYMBOL					= 0x02;
	final public const CHARSET_APPLE_ROMAN				= 0x4D;
	final public const CHARSET_ANSI_JAPANESE_SHIFTJIS	= 0x80;
	final public const CHARSET_ANSI_KOREAN_HANGUL		= 0x81;
	final public const CHARSET_ANSI_KOREAN_JOHAB			= 0x82;
	final public const CHARSET_ANSI_CHINESE_SIMIPLIFIED	= 0x86;		//	gb2312
	final public const CHARSET_ANSI_CHINESE_TRADITIONAL	= 0x88;		//	big5
	final public const CHARSET_ANSI_GREEK				= 0xA1;
	final public const CHARSET_ANSI_TURKISH				= 0xA2;
	final public const CHARSET_ANSI_VIETNAMESE			= 0xA3;
	final public const CHARSET_ANSI_HEBREW				= 0xB1;
	final public const CHARSET_ANSI_ARABIC				= 0xB2;
	final public const CHARSET_ANSI_BALTIC				= 0xBA;
	final public const CHARSET_ANSI_CYRILLIC				= 0xCC;
	final public const CHARSET_ANSI_THAI					= 0xDD;
	final public const CHARSET_ANSI_LATIN_II				= 0xEE;
	final public const CHARSET_OEM_LATIN_I				= 0xFF;

	//  XXX: Constants created!
	/** Font filenames */
	final public const ARIAL								= 'arial.ttf';
	final public const ARIAL_BOLD						= 'arialbd.ttf';
	final public const ARIAL_ITALIC						= 'ariali.ttf';
	final public const ARIAL_BOLD_ITALIC					= 'arialbi.ttf';

	final public const CALIBRI							= 'CALIBRI.TTF';
	final public const CALIBRI_BOLD						= 'CALIBRIB.TTF';
	final public const CALIBRI_ITALIC					= 'CALIBRII.TTF';
	final public const CALIBRI_BOLD_ITALIC				= 'CALIBRIZ.TTF';

	final public const COMIC_SANS_MS						= 'comic.ttf';
	final public const COMIC_SANS_MS_BOLD				= 'comicbd.ttf';

	final public const COURIER_NEW						= 'cour.ttf';
	final public const COURIER_NEW_BOLD					= 'courbd.ttf';
	final public const COURIER_NEW_ITALIC				= 'couri.ttf';
	final public const COURIER_NEW_BOLD_ITALIC			= 'courbi.ttf';

	final public const GEORGIA							= 'georgia.ttf';
	final public const GEORGIA_BOLD						= 'georgiab.ttf';
	final public const GEORGIA_ITALIC					= 'georgiai.ttf';
	final public const GEORGIA_BOLD_ITALIC				= 'georgiaz.ttf';

	final public const IMPACT							= 'impact.ttf';

	final public const LIBERATION_SANS					= 'LiberationSans-Regular.ttf';
	final public const LIBERATION_SANS_BOLD				= 'LiberationSans-Bold.ttf';
	final public const LIBERATION_SANS_ITALIC			= 'LiberationSans-Italic.ttf';
	final public const LIBERATION_SANS_BOLD_ITALIC		= 'LiberationSans-BoldItalic.ttf';

	final public const LUCIDA_CONSOLE					= 'lucon.ttf';
	final public const LUCIDA_SANS_UNICODE				= 'l_10646.ttf';

	final public const MICROSOFT_SANS_SERIF				= 'micross.ttf';

	final public const PALATINO_LINOTYPE					= 'pala.ttf';
	final public const PALATINO_LINOTYPE_BOLD			= 'palab.ttf';
	final public const PALATINO_LINOTYPE_ITALIC			= 'palai.ttf';
	final public const PALATINO_LINOTYPE_BOLD_ITALIC		= 'palabi.ttf';

	final public const SYMBOL							= 'symbol.ttf';

	final public const TAHOMA							= 'tahoma.ttf';
	final public const TAHOMA_BOLD						= 'tahomabd.ttf';

	final public const TIMES_NEW_ROMAN					= 'times.ttf';
	final public const TIMES_NEW_ROMAN_BOLD				= 'timesbd.ttf';
	final public const TIMES_NEW_ROMAN_ITALIC			= 'timesi.ttf';
	final public const TIMES_NEW_ROMAN_BOLD_ITALIC		= 'timesbi.ttf';

	final public const TREBUCHET_MS						= 'trebuc.ttf';
	final public const TREBUCHET_MS_BOLD					= 'trebucbd.ttf';
	final public const TREBUCHET_MS_ITALIC				= 'trebucit.ttf';
	final public const TREBUCHET_MS_BOLD_ITALIC			= 'trebucbi.ttf';

	final public const VERDANA							= 'verdana.ttf';
	final public const VERDANA_BOLD						= 'verdanab.ttf';
	final public const VERDANA_ITALIC					= 'verdanai.ttf';
	final public const VERDANA_BOLD_ITALIC				= 'verdanaz.ttf';

	/**
  * AutoSize method
  */
 private static string $autoSizeMethod = self::AUTOSIZE_METHOD_APPROX;

	/**
	 * Path to folder containing TrueType font .ttf files
	 *
	 * @var string
	 */
	private static $trueTypeFontPath = null;

	/**
	 * How wide is a default column for a given default font and size?
	 * Empirical data found by inspecting real Excel files and reading off the pixel width
	 * in Microsoft Office Excel 2007.
	 *
	 * @var array
	 */
	public static $defaultColumnWidths = ['Arial' => [1 => ['px' => 24, 'width' => 12.00000000], 2 => ['px' => 24, 'width' => 12.00000000], 3 => ['px' => 32, 'width' => 10.66406250], 4 => ['px' => 32, 'width' => 10.66406250], 5 => ['px' => 40, 'width' => 10.00000000], 6 => ['px' => 48, 'width' =>  9.59765625], 7 => ['px' => 48, 'width' =>  9.59765625], 8 => ['px' => 56, 'width' =>  9.33203125], 9 => ['px' => 64, 'width' =>  9.14062500], 10 => ['px' => 64, 'width' =>  9.14062500]], 'Calibri' => [1 => ['px' => 24, 'width' => 12.00000000], 2 => ['px' => 24, 'width' => 12.00000000], 3 => ['px' => 32, 'width' => 10.66406250], 4 => ['px' => 32, 'width' => 10.66406250], 5 => ['px' => 40, 'width' => 10.00000000], 6 => ['px' => 48, 'width' =>  9.59765625], 7 => ['px' => 48, 'width' =>  9.59765625], 8 => ['px' => 56, 'width' =>  9.33203125], 9 => ['px' => 56, 'width' =>  9.33203125], 10 => ['px' => 64, 'width' =>  9.14062500], 11 => ['px' => 64, 'width' =>  9.14062500]], 'Verdana' => [1 => ['px' => 24, 'width' => 12.00000000], 2 => ['px' => 24, 'width' => 12.00000000], 3 => ['px' => 32, 'width' => 10.66406250], 4 => ['px' => 32, 'width' => 10.66406250], 5 => ['px' => 40, 'width' => 10.00000000], 6 => ['px' => 48, 'width' =>  9.59765625], 7 => ['px' => 48, 'width' =>  9.59765625], 8 => ['px' => 64, 'width' =>  9.14062500], 9 => ['px' => 72, 'width' =>  9.00000000], 10 => ['px' => 72, 'width' =>  9.00000000]]];

	/**
	 * Set autoSize method
	 *
	 * @param string $pValue
	 * @return	 boolean					Success or failure
	 */
	public static function setAutoSizeMethod($pValue = self::AUTOSIZE_METHOD_APPROX)
	{
		if (!in_array($pValue,self::$_autoSizeMethods)) {
			return FALSE;
		}

		self::$autoSizeMethod = $pValue;

		return TRUE;
	}

	/**
	 * Get autoSize method
	 *
	 * @return string
	 */
	public static function getAutoSizeMethod()
	{
		return self::$autoSizeMethod;
	}

	/**
	 * Set the path to the folder containing .ttf files. There should be a trailing slash.
	 * Typical locations on variout some platforms:
	 *	<ul>
	 *		<li>C:/Windows/Fonts/</li>
	 *		<li>/usr/share/fonts/truetype/</li>
	 *		<li>~/.fonts/</li>
	 *	</ul>
	 *
	 * @param string $pValue
	 */
	public static function setTrueTypeFontPath($pValue = '')
	{
		self::$trueTypeFontPath = $pValue;
	}

	/**
	 * Get the path to the folder containing .ttf files.
	 *
	 * @return string
	 */
	public static function getTrueTypeFontPath()
	{
		return self::$trueTypeFontPath;
	}

	/**
	 * Calculate an (approximate) OpenXML column width, based on font size and text contained
	 *
	 * @param 	PHPExcel_Style_Font			$font			Font object
	 * @param 	PHPExcel_RichText|string	$cellText		Text to calculate width
	 * @param 	integer						$rotation		Rotation angle
	 * @param 	PHPExcel_Style_Font|NULL	$defaultFont	Font object
	 * @return 	integer		Column width
	 */
	public static function calculateColumnWidth(PHPExcel_Style_Font $font, $cellText = '', $rotation = 0, PHPExcel_Style_Font $defaultFont = null) {

		// If it is rich text, use plain text
		if ($cellText instanceof PHPExcel_RichText) {
			$cellText = $cellText->getPlainText();
		}

		// Special case if there are one or more newline characters ("\n")
		if (str_contains($cellText, "\n")) {
			$lineTexts = explode("\n", $cellText);
			$lineWitdhs = [];
			foreach ($lineTexts as $lineText) {
				$lineWidths[] = self::calculateColumnWidth($font, $lineText, $rotation = 0, $defaultFont);
			}
			return max($lineWidths); // width of longest line in cell
		}

		// Try to get the exact text width in pixels
		try {
			// If autosize method is set to 'approx', use approximation
			if (self::$autoSizeMethod == self::AUTOSIZE_METHOD_APPROX) {
				throw new PHPExcel_Exception('AutoSize method is set to approx');
			}

			// Width of text in pixels excl. padding
			$columnWidth = self::getTextWidthPixelsExact($cellText, $font, $rotation);

			// Excel adds some padding, use 1.07 of the width of an 'n' glyph
			$columnWidth += ceil(self::getTextWidthPixelsExact('0', $font, 0) * 1.07); // pixels incl. padding

		} catch (PHPExcel_Exception) {
			// Width of text in pixels excl. padding, approximation
			$columnWidth = self::getTextWidthPixelsApprox($cellText, $font, $rotation);

			// Excel adds some padding, just use approx width of 'n' glyph
			$columnWidth += self::getTextWidthPixelsApprox('n', $font, 0);
		}

		// Convert from pixel width to column width
		$columnWidth = PHPExcel_Shared_Drawing::pixelsToCellDimension($columnWidth, $defaultFont);

		// Return
		return round($columnWidth, 6);
	}

	/**
	 * Get GD text width in pixels for a string of text in a certain font at a certain rotation angle
	 *
	 * @param string $text
	 * @param PHPExcel_Style_Font
	 * @param int $rotation
	 * @return int
	 * @throws PHPExcel_Exception
	 */
	public static function getTextWidthPixelsExact($text, PHPExcel_Style_Font $font, $rotation = 0) {
		if (!function_exists('imagettfbbox')) {
			throw new PHPExcel_Exception('GD library needs to be enabled');
		}

		// font size should really be supplied in pixels in GD2,
		// but since GD2 seems to assume 72dpi, pixels and points are the same
		$fontFile = self::getTrueTypeFontFileFromFont($font);
		$textBox = imagettfbbox($font->getSize(), $rotation, $fontFile, $text);

		// Get corners positions
		$lowerLeftCornerX  = $textBox[0];
		$lowerLeftCornerY  = $textBox[1];
		$lowerRightCornerX = $textBox[2];
		$lowerRightCornerY = $textBox[3];
		$upperRightCornerX = $textBox[4];
		$upperRightCornerY = $textBox[5];
		$upperLeftCornerX  = $textBox[6];
		$upperLeftCornerY  = $textBox[7];

		// Consider the rotation when calculating the width
		$textWidth = max($lowerRightCornerX - $upperLeftCornerX, $upperRightCornerX - $lowerLeftCornerX);

		return $textWidth;
	}

	/**
  * Get approximate width in pixels for a string of text in a certain font at a certain rotation angle
  *
  * @param string $columnText
  * @param int $rotation
  * @return int Text width in pixels (no padding added)
  */
 public static function getTextWidthPixelsApprox($columnText, PHPExcel_Style_Font $font = null, $rotation = 0)
	{
		$fontName = $font->getName();
		$fontSize = $font->getSize();

		// Calculate column width in pixels. We assume fixed glyph width. Result varies with font name and size.
		switch ($fontName) {
			case 'Calibri':
				// value 8.26 was found via interpolation by inspecting real Excel files with Calibri 11 font.
				$columnWidth = (int) (8.26 * PHPExcel_Shared_String::CountCharacters($columnText));
				$columnWidth = $columnWidth * $fontSize / 11; // extrapolate from font size
				break;

			case 'Arial':
				// value 7 was found via interpolation by inspecting real Excel files with Arial 10 font.
				$columnWidth = (int) (7 * PHPExcel_Shared_String::CountCharacters($columnText));
				$columnWidth = $columnWidth * $fontSize / 10; // extrapolate from font size
				break;

			case 'Verdana':
				// value 8 was found via interpolation by inspecting real Excel files with Verdana 10 font.
				$columnWidth = (int) (8 * PHPExcel_Shared_String::CountCharacters($columnText));
				$columnWidth = $columnWidth * $fontSize / 10; // extrapolate from font size
				break;

			default:
				// just assume Calibri
				$columnWidth = (int) (8.26 * PHPExcel_Shared_String::CountCharacters($columnText));
				$columnWidth = $columnWidth * $fontSize / 11; // extrapolate from font size
				break;
		}

		// Calculate approximate rotated column width
		if ($rotation !== 0) {
			if ($rotation == -165) {
				// stacked text
				$columnWidth = 4; // approximation
			} else {
				// rotated text
				$columnWidth = $columnWidth * cos(deg2rad($rotation))
								+ $fontSize * abs(sin(deg2rad($rotation))) / 5; // approximation
			}
		}

		// pixel width is an integer
		$columnWidth = (int) $columnWidth;
		return $columnWidth;
	}

	/**
	 * Calculate an (approximate) pixel size, based on a font points size
	 *
	 * @param 	int		$fontSizeInPoints	Font size (in points)
	 * @return 	int		Font size (in pixels)
	 */
	public static function fontSizeToPixels($fontSizeInPoints = 11) {
		return (int) ((4 / 3) * $fontSizeInPoints);
	}

	/**
	 * Calculate an (approximate) pixel size, based on inch size
	 *
	 * @param 	int		$sizeInInch	Font size (in inch)
	 * @return 	int		Size (in pixels)
	 */
	public static function inchSizeToPixels($sizeInInch = 1) {
		return ($sizeInInch * 96);
	}

	/**
	 * Calculate an (approximate) pixel size, based on centimeter size
	 *
	 * @param 	int		$sizeInCm	Font size (in centimeters)
	 * @return 	int		Size (in pixels)
	 */
	public static function centimeterSizeToPixels($sizeInCm = 1) {
		return ($sizeInCm * 37.795275591);
	}

	/**
	 * Returns the font path given the font
	 *
	 * @param PHPExcel_Style_Font
	 * @return string Path to TrueType font file
	 */
	public static function getTrueTypeFontFileFromFont($font) {
		if (!file_exists(self::$trueTypeFontPath) || !is_dir(self::$trueTypeFontPath)) {
			throw new PHPExcel_Exception('Valid directory to TrueType Font files not specified');
		}

		$name		= $font->getName();
		$bold		= $font->getBold();
		$italic		= $font->getItalic();

		// Check if we can map font to true type font file
		$fontFile = match ($name) {
      'Arial' => $bold ? ($italic ? self::ARIAL_BOLD_ITALIC : self::ARIAL_BOLD)
 						  : ($italic ? self::ARIAL_ITALIC : self::ARIAL),
      'Calibri' => $bold ? ($italic ? self::CALIBRI_BOLD_ITALIC : self::CALIBRI_BOLD)
 						  : ($italic ? self::CALIBRI_ITALIC : self::CALIBRI),
      'Courier New' => $bold ? ($italic ? self::COURIER_NEW_BOLD_ITALIC : self::COURIER_NEW_BOLD)
 						  : ($italic ? self::COURIER_NEW_ITALIC : self::COURIER_NEW),
      'Comic Sans MS' => $bold ? self::COMIC_SANS_MS_BOLD : self::COMIC_SANS_MS,
      'Georgia' => $bold ? ($italic ? self::GEORGIA_BOLD_ITALIC : self::GEORGIA_BOLD)
 						  : ($italic ? self::GEORGIA_ITALIC : self::GEORGIA),
      'Impact' => self::IMPACT,
      'Liberation Sans' => $bold ? ($italic ? self::LIBERATION_SANS_BOLD_ITALIC : self::LIBERATION_SANS_BOLD)
 						  : ($italic ? self::LIBERATION_SANS_ITALIC : self::LIBERATION_SANS),
      'Lucida Console' => self::LUCIDA_CONSOLE,
      'Lucida Sans Unicode' => self::LUCIDA_SANS_UNICODE,
      'Microsoft Sans Serif' => self::MICROSOFT_SANS_SERIF,
      'Palatino Linotype' => $bold ? ($italic ? self::PALATINO_LINOTYPE_BOLD_ITALIC : self::PALATINO_LINOTYPE_BOLD)
 						  : ($italic ? self::PALATINO_LINOTYPE_ITALIC : self::PALATINO_LINOTYPE),
      'Symbol' => self::SYMBOL,
      'Tahoma' => $bold ? self::TAHOMA_BOLD : self::TAHOMA,
      'Times New Roman' => $bold ? ($italic ? self::TIMES_NEW_ROMAN_BOLD_ITALIC : self::TIMES_NEW_ROMAN_BOLD)
 						  : ($italic ? self::TIMES_NEW_ROMAN_ITALIC : self::TIMES_NEW_ROMAN),
      'Trebuchet MS' => $bold ? ($italic ? self::TREBUCHET_MS_BOLD_ITALIC : self::TREBUCHET_MS_BOLD)
 						  : ($italic ? self::TREBUCHET_MS_ITALIC : self::TREBUCHET_MS),
      'Verdana' => $bold ? ($italic ? self::VERDANA_BOLD_ITALIC : self::VERDANA_BOLD)
 						  : ($italic ? self::VERDANA_ITALIC : self::VERDANA),
      default => throw new PHPExcel_Exception('Unknown font name "'. $name .'". Cannot map to TrueType font file'),
  };

		$fontFile = self::$trueTypeFontPath . $fontFile;

		// Check if file actually exists
		if (!file_exists($fontFile)) {
			throw New PHPExcel_Exception('TrueType Font file not found');
		}

		return $fontFile;
	}

	/**
	 * Returns the associated charset for the font name.
	 *
	 * @param string $name Font name
	 * @return int Character set code
	 */
	public static function getCharsetFromFontName($name)
	{
		return match ($name) {
      'EucrosiaUPC' => self::CHARSET_ANSI_THAI,
      'Wingdings' => self::CHARSET_SYMBOL,
      'Wingdings 2' => self::CHARSET_SYMBOL,
      'Wingdings 3' => self::CHARSET_SYMBOL,
      default => self::CHARSET_ANSI_LATIN,
  };
	}

	/**
	 * Get the effective column width for columns without a column dimension or column with width -1
	 * For example, for Calibri 11 this is 9.140625 (64 px)
	 *
	 * @param PHPExcel_Style_Font $font The workbooks default font
	 * @param boolean $pPixels true = return column width in pixels, false = return in OOXML units
	 * @return mixed Column width
	 */
	public static function getDefaultColumnWidthByFont(PHPExcel_Style_Font $font, $pPixels = false)
	{
		if (isset(self::$defaultColumnWidths[$font->getName()][$font->getSize()])) {
			// Exact width can be determined
			$columnWidth = $pPixels ?
				self::$defaultColumnWidths[$font->getName()][$font->getSize()]['px']
					: self::$defaultColumnWidths[$font->getName()][$font->getSize()]['width'];

		} else {
			// We don't have data for this particular font and size, use approximation by
			// extrapolating from Calibri 11
			$columnWidth = $pPixels ?
				self::$defaultColumnWidths['Calibri'][11]['px']
					: self::$defaultColumnWidths['Calibri'][11]['width'];
			$columnWidth = $columnWidth * $font->getSize() / 11;

			// Round pixels to closest integer
			if ($pPixels) {
				$columnWidth = (int) round($columnWidth);
			}
		}

		return $columnWidth;
	}

	/**
	 * Get the effective row height for rows without a row dimension or rows with height -1
	 * For example, for Calibri 11 this is 15 points
	 *
	 * @param PHPExcel_Style_Font $font The workbooks default font
	 * @return float Row height in points
	 */
	public static function getDefaultRowHeightByFont(PHPExcel_Style_Font $font)
	{
		$rowHeight = match ($font->getName()) {
      'Arial' => match ($font->getSize()) {
          10 => 12.75,
          9 => 12,
          8 => 11.25,
          7 => 9,
          6, 5 => 8.25,
          4 => 6.75,
          3 => 6,
          2, 1 => 5.25,
          default => 12.75 * $font->getSize() / 10,
      },
      'Calibri' => match ($font->getSize()) {
          11 => 15,
          10 => 12.75,
          9 => 12,
          8 => 11.25,
          7 => 9,
          6, 5 => 8.25,
          4 => 6.75,
          3 => 6.00,
          2, 1 => 5.25,
          default => 15 * $font->getSize() / 11,
      },
      'Verdana' => match ($font->getSize()) {
          10 => 12.75,
          9 => 11.25,
          8 => 10.50,
          7 => 9.00,
          6, 5 => 8.25,
          4 => 6.75,
          3 => 6,
          2, 1 => 5.25,
          default => 12.75 * $font->getSize() / 10,
      },
      default => 15 * $font->getSize() / 11,
  };

		return $rowHeight;
	}

}
