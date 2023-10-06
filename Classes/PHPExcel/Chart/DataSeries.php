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
 * @version	##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Chart_DataSeries
 *
 * @category	PHPExcel
 * @package		PHPExcel_Chart
 * @copyright	Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Chart_DataSeries
{

	final public const TYPE_BARCHART			= 'barChart';
	final public const TYPE_BARCHART_3D		= 'bar3DChart';
	final public const TYPE_LINECHART		= 'lineChart';
	final public const TYPE_LINECHART_3D		= 'line3DChart';
	final public const TYPE_AREACHART		= 'areaChart';
	final public const TYPE_AREACHART_3D		= 'area3DChart';
	final public const TYPE_PIECHART			= 'pieChart';
	final public const TYPE_PIECHART_3D		= 'pie3DChart';
	final public const TYPE_DOUGHTNUTCHART	= 'doughnutChart';
	final public const TYPE_DONUTCHART		= self::TYPE_DOUGHTNUTCHART;	//	Synonym
	final public const TYPE_SCATTERCHART		= 'scatterChart';
	final public const TYPE_SURFACECHART		= 'surfaceChart';
	final public const TYPE_SURFACECHART_3D	= 'surface3DChart';
	final public const TYPE_RADARCHART		= 'radarChart';
	final public const TYPE_BUBBLECHART		= 'bubbleChart';
	final public const TYPE_STOCKCHART		= 'stockChart';

	final public const GROUPING_CLUSTERED			= 'clustered';
	final public const GROUPING_STACKED				= 'stacked';
	final public const GROUPING_PERCENT_STACKED		= 'percentStacked';
	final public const GROUPING_STANDARD				= 'standard';

	final public const DIRECTION_BAR			= 'bar';
	final public const DIRECTION_HORIZONTAL	= self::DIRECTION_BAR;
	final public const DIRECTION_COL			= 'col';
	final public const DIRECTION_COLUMN		= self::DIRECTION_COL;
	final public const DIRECTION_VERTICAL	= self::DIRECTION_COL;

	final public const STYLE_LINEMARKER		= 'lineMarker';
	final public const STYLE_SMOOTHMARKER	= 'smoothMarker';
	final public const STYLE_MARKER			= 'marker';
	final public const STYLE_FILLED			= 'filled';

	/**
	 * Plot Direction
	 *
	 * @var boolean
	 */
	private $_plotDirection = null;

	/**
	 * Plot Label
	 *
	 * @var array of PHPExcel_Chart_DataSeriesValues
	 */
	private array $_plotLabel = [];

	/**
	 * Plot Category
	 *
	 * @var array of PHPExcel_Chart_DataSeriesValues
	 */
	private array $_plotCategory = [];

	/**
	 * Plot Values
	 *
	 * @var array of PHPExcel_Chart_DataSeriesValues
	 */
	private array $_plotValues = [];

	/**
  * Create a new PHPExcel_Chart_DataSeries
  * @param string $plotType
  * @param bool $plotGrouping
  * @param string $plotStyle
  * @param mixed[] $plotOrder
  * @param string $smoothLine
  */
 public function __construct(/**
  * Series Plot Type
  */
 private $_plotType = null, /**
  * Plot Grouping Type
  */
 private $_plotGrouping = null, /**
  * Order of plots in Series
  *
  * @var array of integer
  */
 private $_plotOrder = [], $plotLabel = [], $plotCategory = [], $plotValues = [], /**
  * Smooth Line
  */
 private $_smoothLine = null, /**
  * Plot Style
  */
 private $_plotStyle = null)
	{
		$keys = array_keys($plotValues);
		$this->_plotValues = $plotValues;
		if ((count($plotLabel) == 0) || (is_null($plotLabel[$keys[0]]))) {
			$plotLabel[$keys[0]] = new PHPExcel_Chart_DataSeriesValues();
		}

		$this->_plotLabel = $plotLabel;
		if ((count($plotCategory) == 0) || (is_null($plotCategory[$keys[0]]))) {
			$plotCategory[$keys[0]] = new PHPExcel_Chart_DataSeriesValues();
		}
		$this->_plotCategory = $plotCategory;
	}

	/**
	 * Get Plot Type
	 *
	 * @return string
	 */
	public function getPlotType() {
		return $this->_plotType;
	}

	/**
	 * Set Plot Type
	 *
	 * @param string $plotType
	 */
	public function setPlotType($plotType = '') {
		$this->_plotType = $plotType;
	}

	/**
	 * Get Plot Grouping Type
	 *
	 * @return string
	 */
	public function getPlotGrouping() {
		return $this->_plotGrouping;
	}

	/**
	 * Set Plot Grouping Type
	 *
	 * @param string $groupingType
	 */
	public function setPlotGrouping($groupingType = null) {
		$this->_plotGrouping = $groupingType;
	}

	/**
	 * Get Plot Direction
	 *
	 * @return string
	 */
	public function getPlotDirection() {
		return $this->_plotDirection;
	}

	/**
	 * Set Plot Direction
	 *
	 * @param string $plotDirection
	 */
	public function setPlotDirection($plotDirection = null) {
		$this->_plotDirection = $plotDirection;
	}

	/**
	 * Get Plot Order
	 *
	 * @return string
	 */
	public function getPlotOrder() {
		return $this->_plotOrder;
	}

	/**
	 * Get Plot Labels
	 *
	 * @return array of PHPExcel_Chart_DataSeriesValues
	 */
	public function getPlotLabels() {
		return $this->_plotLabel;
	}

	/**
	 * Get Plot Label by Index
	 *
	 * @return PHPExcel_Chart_DataSeriesValues
	 */
	public function getPlotLabelByIndex($index) {
		$keys = array_keys($this->_plotLabel);
		if (in_array($index,$keys)) {
			return $this->_plotLabel[$index];
		} elseif(isset($keys[$index])) {
			return $this->_plotLabel[$keys[$index]];
		}
		return false;
	}

	/**
	 * Get Plot Categories
	 *
	 * @return array of PHPExcel_Chart_DataSeriesValues
	 */
	public function getPlotCategories() {
		return $this->_plotCategory;
	}

	/**
	 * Get Plot Category by Index
	 *
	 * @return PHPExcel_Chart_DataSeriesValues
	 */
	public function getPlotCategoryByIndex($index) {
		$keys = array_keys($this->_plotCategory);
		if (in_array($index,$keys)) {
			return $this->_plotCategory[$index];
		} elseif(isset($keys[$index])) {
			return $this->_plotCategory[$keys[$index]];
		}
		return false;
	}

	/**
	 * Get Plot Style
	 *
	 * @return string
	 */
	public function getPlotStyle() {
		return $this->_plotStyle;
	}

	/**
	 * Set Plot Style
	 *
	 * @param string $plotStyle
	 */
	public function setPlotStyle($plotStyle = null) {
		$this->_plotStyle = $plotStyle;
	}

	/**
	 * Get Plot Values
	 *
	 * @return array of PHPExcel_Chart_DataSeriesValues
	 */
	public function getPlotValues() {
		return $this->_plotValues;
	}

	/**
	 * Get Plot Values by Index
	 *
	 * @return PHPExcel_Chart_DataSeriesValues
	 */
	public function getPlotValuesByIndex($index) {
		$keys = array_keys($this->_plotValues);
		if (in_array($index,$keys)) {
			return $this->_plotValues[$index];
		} elseif(isset($keys[$index])) {
			return $this->_plotValues[$keys[$index]];
		}
		return false;
	}

	/**
	 * Get Number of Plot Series
	 *
	 * @return integer
	 */
	public function getPlotSeriesCount() {
		return count($this->_plotValues);
	}

	/**
	 * Get Smooth Line
	 *
	 * @return boolean
	 */
	public function getSmoothLine() {
		return $this->_smoothLine;
	}

	/**
	 * Set Smooth Line
	 *
	 * @param boolean $smoothLine
	 */
	public function setSmoothLine($smoothLine = TRUE) {
		$this->_smoothLine = $smoothLine;
	}

	public function refresh(PHPExcel_Worksheet $worksheet) {
	    foreach($this->_plotValues as $plotValues) {
			if ($plotValues !== NULL)
				$plotValues->refresh($worksheet, TRUE);
		}
		foreach($this->_plotLabel as $plotValues) {
			if ($plotValues !== NULL)
				$plotValues->refresh($worksheet, TRUE);
		}
		foreach($this->_plotCategory as $plotValues) {
			if ($plotValues !== NULL)
				$plotValues->refresh($worksheet, FALSE);
		}
	}

}
