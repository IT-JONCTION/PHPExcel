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
 * @category	PHPExcel
 * @package		PHPExcel_Chart
 * @copyright	Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license		http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version		##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Chart
 *
 * @category	PHPExcel
 * @package		PHPExcel_Chart
 * @copyright	Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Chart
{
	/**
  * Worksheet
  */
 private ?\PHPExcel_Worksheet $_worksheet = null;


	/**
  * Top-Left Cell Position
  */
 private string $_topLeftCellRef = 'A1';


	/**
  * Top-Left X-Offset
  */
 private int $_topLeftXOffset = 0;


	/**
  * Top-Left Y-Offset
  */
 private int $_topLeftYOffset = 0;


	/**
  * Bottom-Right Cell Position
  */
 private string $_bottomRightCellRef = 'A1';


	/**
  * Bottom-Right X-Offset
  */
 private int $_bottomRightXOffset = 10;


	/**
  * Bottom-Right Y-Offset
  */
 private int $_bottomRightYOffset = 10;


	/**
  * Create a new PHPExcel_Chart
  * @param string $name
  * @param bool $plotVisibleOnly
  * @param string $displayBlanksAs
  */
 public function __construct(
     /**
      * Chart Name
      */
     private $_name,
     /**
      * Chart Title
      */
     private ?\PHPExcel_Chart_Title $_title = null,
     /**
      * Chart Legend
      */
     private ?\PHPExcel_Chart_Legend $_legend = null,
     /**
      * Chart Plot Area
      */
     private readonly ?\PHPExcel_Chart_PlotArea $_plotArea = null,
     /**
      * Plot Visible Only
      */
     private $_plotVisibleOnly = true,
     /**
      * Display Blanks as
      */
     private $_displayBlanksAs = '0',
     /**
      * X-Axis Label
      */
     private ?\PHPExcel_Chart_Title $_xAxisLabel = null,
     /**
      * Y-Axis Label
      */
     private ?\PHPExcel_Chart_Title $_yAxisLabel = null
 )
 {
 }

	/**
	 * Get Name
	 *
	 * @return string
	 */
	public function getName() {
		return $this->_name;
	}

	/**
	 * Get Worksheet
	 *
	 * @return PHPExcel_Worksheet
	 */
	public function getWorksheet() {
		return $this->_worksheet;
	}

	/**
  * Set Worksheet
  *
  * @throws	PHPExcel_Chart_Exception
  * @return PHPExcel_Chart
  */
 public function setWorksheet(PHPExcel_Worksheet $pValue = null) {
		$this->_worksheet = $pValue;

		return $this;
	}

	/**
	 * Get Title
	 *
	 * @return PHPExcel_Chart_Title
	 */
	public function getTitle() {
		return $this->_title;
	}

	/**
  * Set Title
  *
  * @return	PHPExcel_Chart
  */
 public function setTitle(PHPExcel_Chart_Title $title) {
		$this->_title = $title;

		return $this;
	}

	/**
	 * Get Legend
	 *
	 * @return PHPExcel_Chart_Legend
	 */
	public function getLegend() {
		return $this->_legend;
	}

	/**
  * Set Legend
  *
  * @return	PHPExcel_Chart
  */
 public function setLegend(PHPExcel_Chart_Legend $legend) {
		$this->_legend = $legend;

		return $this;
	}

	/**
	 * Get X-Axis Label
	 *
	 * @return PHPExcel_Chart_Title
	 */
	public function getXAxisLabel() {
		return $this->_xAxisLabel;
	}

	/**
  * Set X-Axis Label
  *
  * @return	PHPExcel_Chart
  */
 public function setXAxisLabel(PHPExcel_Chart_Title $label) {
		$this->_xAxisLabel = $label;

		return $this;
	}

	/**
	 * Get Y-Axis Label
	 *
	 * @return PHPExcel_Chart_Title
	 */
	public function getYAxisLabel() {
		return $this->_yAxisLabel;
	}

	/**
  * Set Y-Axis Label
  *
  * @return	PHPExcel_Chart
  */
 public function setYAxisLabel(PHPExcel_Chart_Title $label) {
		$this->_yAxisLabel = $label;

		return $this;
	}

	/**
	 * Get Plot Area
	 *
	 * @return PHPExcel_Chart_PlotArea
	 */
	public function getPlotArea() {
		return $this->_plotArea;
	}

	/**
	 * Get Plot Visible Only
	 *
	 * @return boolean
	 */
	public function getPlotVisibleOnly() {
		return $this->_plotVisibleOnly;
	}

	/**
	 * Set Plot Visible Only
	 *
	 * @param boolean $plotVisibleOnly
	 * @return PHPExcel_Chart
	 */
	public function setPlotVisibleOnly($plotVisibleOnly = true) {
		$this->_plotVisibleOnly = $plotVisibleOnly;

		return $this;
	}

	/**
	 * Get Display Blanks as
	 *
	 * @return string
	 */
	public function getDisplayBlanksAs() {
		return $this->_displayBlanksAs;
	}

	/**
	 * Set Display Blanks as
	 *
	 * @param string $displayBlanksAs
	 * @return PHPExcel_Chart
	 */
	public function setDisplayBlanksAs($displayBlanksAs = '0') {
		$this->_displayBlanksAs = $displayBlanksAs;
	}


	/**
	 * Set the Top Left position for the chart
	 *
	 * @param	string	$cell
	 * @param	integer	$xOffset
	 * @param	integer	$yOffset
	 * @return PHPExcel_Chart
	 */
	public function setTopLeftPosition($cell, $xOffset=null, $yOffset=null) {
		$this->_topLeftCellRef = $cell;
		if (!is_null($xOffset))
			$this->setTopLeftXOffset($xOffset);
		if (!is_null($yOffset))
			$this->setTopLeftYOffset($yOffset);

		return $this;
	}

	/**
	 * Get the top left position of the chart
	 *
	 * @return array	an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
	 */
	public function getTopLeftPosition() {
		return ['cell'	=> $this->_topLeftCellRef, 'xOffset'	=> $this->_topLeftXOffset, 'yOffset'	=> $this->_topLeftYOffset];
	}

	/**
	 * Get the cell address where the top left of the chart is fixed
	 *
	 * @return string
	 */
	public function getTopLeftCell() {
		return $this->_topLeftCellRef;
	}

	/**
	 * Set the Top Left cell position for the chart
	 *
	 * @param	string	$cell
	 * @return PHPExcel_Chart
	 */
	public function setTopLeftCell($cell) {
		$this->_topLeftCellRef = $cell;

		return $this;
	}

	/**
	 * Set the offset position within the Top Left cell for the chart
	 *
	 * @param	integer	$xOffset
	 * @param	integer	$yOffset
	 * @return PHPExcel_Chart
	 */
	public function setTopLeftOffset($xOffset=null,$yOffset=null) {
		if (!is_null($xOffset))
			$this->setTopLeftXOffset($xOffset);
		if (!is_null($yOffset))
			$this->setTopLeftYOffset($yOffset);

		return $this;
	}

	/**
	 * Get the offset position within the Top Left cell for the chart
	 *
	 * @return integer[]
	 */
	public function getTopLeftOffset() {
		return ['X' => $this->_topLeftXOffset, 'Y' => $this->_topLeftYOffset];
	}

	public function setTopLeftXOffset($xOffset) {
		$this->_topLeftXOffset = $xOffset;

		return $this;
	}

	public function getTopLeftXOffset() {
		return $this->_topLeftXOffset;
	}

	public function setTopLeftYOffset($yOffset) {
		$this->_topLeftYOffset = $yOffset;

		return $this;
	}

	public function getTopLeftYOffset() {
		return $this->_topLeftYOffset;
	}

	/**
	 * Set the Bottom Right position of the chart
	 *
	 * @param	string	$cell
	 * @param	integer	$xOffset
	 * @param	integer	$yOffset
	 * @return PHPExcel_Chart
	 */
	public function setBottomRightPosition($cell, $xOffset=null, $yOffset=null) {
		$this->_bottomRightCellRef = $cell;
		if (!is_null($xOffset))
			$this->setBottomRightXOffset($xOffset);
		if (!is_null($yOffset))
			$this->setBottomRightYOffset($yOffset);

		return $this;
	}

	/**
	 * Get the bottom right position of the chart
	 *
	 * @return array	an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
	 */
	public function getBottomRightPosition() {
		return ['cell'	=> $this->_bottomRightCellRef, 'xOffset'	=> $this->_bottomRightXOffset, 'yOffset'	=> $this->_bottomRightYOffset];
	}

	public function setBottomRightCell($cell) {
		$this->_bottomRightCellRef = $cell;

		return $this;
	}

	/**
	 * Get the cell address where the bottom right of the chart is fixed
	 *
	 * @return string
	 */
	public function getBottomRightCell() {
		return $this->_bottomRightCellRef;
	}

	/**
	 * Set the offset position within the Bottom Right cell for the chart
	 *
	 * @param	integer	$xOffset
	 * @param	integer	$yOffset
	 * @return PHPExcel_Chart
	 */
	public function setBottomRightOffset($xOffset=null,$yOffset=null) {
		if (!is_null($xOffset))
			$this->setBottomRightXOffset($xOffset);
		if (!is_null($yOffset))
			$this->setBottomRightYOffset($yOffset);

		return $this;
	}

	/**
	 * Get the offset position within the Bottom Right cell for the chart
	 *
	 * @return integer[]
	 */
	public function getBottomRightOffset() {
		return ['X' => $this->_bottomRightXOffset, 'Y' => $this->_bottomRightYOffset];
	}

	public function setBottomRightXOffset($xOffset) {
		$this->_bottomRightXOffset = $xOffset;

		return $this;
	}

	public function getBottomRightXOffset() {
		return $this->_bottomRightXOffset;
	}

	public function setBottomRightYOffset($yOffset) {
		$this->_bottomRightYOffset = $yOffset;

		return $this;
	}

	public function getBottomRightYOffset() {
		return $this->_bottomRightYOffset;
	}


	public function refresh() {
		if ($this->_worksheet !== NULL) {
			$this->_plotArea->refresh($this->_worksheet);
		}
	}

	public function render($outputDestination = null) {
		$libraryName = PHPExcel_Settings::getChartRendererName();
		if (is_null($libraryName)) {
			return false;
		}
		//	Ensure that data series values are up-to-date before we render
		$this->refresh();

		$libraryPath = PHPExcel_Settings::getChartRendererPath();
		$includePath = str_replace('\\','/',get_include_path());
		$rendererPath = str_replace('\\','/',$libraryPath);
		if (!str_contains($rendererPath,$includePath)) {
			set_include_path(get_include_path() . PATH_SEPARATOR . $libraryPath);
		}

		$rendererName = 'PHPExcel_Chart_Renderer_'.$libraryName;
		$renderer = new $rendererName($this);

		if ($outputDestination == 'php://output') {
			$outputDestination = null;
		}
		return $renderer->render($outputDestination);
	}

}
