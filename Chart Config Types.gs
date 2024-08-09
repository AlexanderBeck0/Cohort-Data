// ######
// These are all taken from Google's documentation for Bar Charts. (https://developers.google.com/apps-script/chart-configuration-options)
// I did not find any intellisense for these options, so I took it upon myself to port it all over into JSDoc.
// They are used in updateChart in Dashboard Utilities, within the options. Enjoy customizing the charts!
// ######
/**
 * @typedef {Object} TextStyle
 * @property {string | null} [color] The color can be any HTML color string, for example: 'red' or '#00cc00'. Set to `undefined` to leave as is, and `null` to remove.
 * @property {string | null} [fontName = '<global-font-name>'] Defaults to `<global-font-name>`
 * @property {number | null} [fontSize = '<global-font-size>'] Defaults to `<global-font-size>`
 * @property {boolean | null} [bold] TextStyle property that specifies if the text is bolded. Set to `undefined` to leave as is, and `null` to remove.
 * @property {boolean | null} [italic] TextStyle property that specifies if the text is italicized. Set to `undefined` to leave as is, and `null` to remove.
 */

/**
 * @typedef {Object} ChartArea An object with members to configure the placement and size of the chart area (where the chart itself is drawn, excluding axis and legends). Two formats are supported: a number, or a number followed by %. A simple number is a value in pixels; a number followed by % is a percentage. Example: 
 * ```javascript
 * chartArea:{left:20,top:0,width:'50%',height:'75%'}
 * ```
 * @property {string | {stroke?: string, strokeWidth?: number}} [backgroundColor = 'white'] Chart area background color. When a string is used, it can be either a hex string (e.g., `'#fdc'`) or an English color name. When an object is used, the following properties can be provided:
 * - `stroke`: The color, provided as a hex string or English color name.
 * - `strokeWidth`: If provided, draws a border around the chart area of the given width (and with the color of `stroke`).
 * @property {number | string} [height = auto] Chart area height.
 * @property {number | string} [left = auto] How far to draw the chart from the left border.
 * @property {number | string} [top = auto] How far to draw the chart from the top border.
 * @property {number | string} [width = auto] Chart area width.
 */

/**
 * @typedef {Object} Gridlines An object with members to configure the gridlines on the axis. Note that axis gridlines are drawn horizontally.
 * @property {string} [color = '#CCC'] The color of the gridlines inside the chart area. Specify a valid HTML color string.
 * @property {number} [count = -1] The approximate number of gridlines inside the chart area. If you specify a positive number for `gridlines.count`, it will be used to compute the `minSpacing` between gridlines. You can specify a value of `1` to only draw one gridline, or `0` to draw no gridlines. Specify `-1`, which is the default, to automatically compute the number of gridlines based on other options.
 */

/**
 * @typedef {Object} MinorGridlines An object with members to configure the minor gridlines on the axis, similar to the gridlines option.
 * @property {string | null} [color] The color of the minor gridlines inside the chart area. Specify a valid HTML color string. Defaults to A blend of the gridline and background colors.
 * @property {number} [count = 1] The `minorGridlines.count` option is mostly deprecated, except for disabling minor gridlines by setting the count to `0`. The number of minor gridlines depends on the interval between major gridlines and the minimum required space.
 */

/**
 * @typedef {Object} ViewWindow Specifies the cropping range of the axis.
 * @property {number | null} [max = auto] The maximum data value to render. Ignored when `viewWindowMode` is 'pretty' or 'maximized'.
 * @property {number | null} [min = auto] The minimum data value to render. Ignored when `viewWindowMode` is 'pretty' or 'maximized'.
 */

/**
 * @typedef {Object} Axis
 * @property {number} [direction = 1] The direction in which the values along the axis grow. By default, low values are on the bottom of the chart. Specify `-1` to reverse the order of the values. **Note**: Only takes in `1` or `-1`.
 * @property {Gridlines | null} [gridlines = null] An object with members to configure the gridlines on the axis. Note that axis gridlines are drawn horizontally.
 * @property {boolean} [logScale = false] If `true`, makes the axis a logarithmic scale. **Note**: All values must be positive.
 * @property {number} [maxValue = automatic] Moves the max value of the axis to the specified value; this will be upward in most charts. Ignored if this is set to a value smaller than the maximum y-value of the data. `viewWindow.max` overrides this property.
 * @property {MinorGridlines | null} [minorGridlines = null] An object with members to configure the minor gridlines on the axis, similar to the gridlines option.
 * @property {number | null} [minValue = null] Moves the min value of the axis to the specified value; this will be downward in most charts. Ignored if this is set to a value greater than the minimum y-value of the data. `viewWindow.min` overrides this property.
 * @property {'out' | 'in' | 'none' | null} [textPosition = 'out'] Position of the axis text, relative to the chart area. Supported values: `'out'`, `'in'`, `'none'`.
 * @property {TextStyle} [textStyle = {color: 'black', fontName: '<global-font-name>', fontSize: '<global-font-size>'}] An object that specifies the axis text style.
 * @property {string | null} [title = null] Specifies a title for the axis. Set to `undefined` to leave as is, and `null` to remove.
 * @property {TextStyle} [titleTextStyle = {color: 'black', fontName: '<global-font-name>', fontSize: '<global-font-size>'}] An object that specifies the axis title text style.
 * @property {ViewWindow | null} [viewWindow = null] Specifies the cropping range of the axis.
 */

/**
 * @typedef {Object} Legend An object with members to configure various aspects of the legend.
 * @property {'bottom' | 'left' | 'in' | 'none' | 'right' | 'top'} [position = 'right'] Position of the legend. Can be one of the following:
 * - `'bottom'` - Below the chart.
 * - `'left'` - To the left of the chart, provided the left axis has no series associated with it. So if you want the legend on the left, use the option `targetAxisIndex: 1`.
 * - `'in'` - Inside the chart, by the top left corner.
 * - `'none'` - No legend is displayed.
 * - `'right'` - To the right of the chart. Incompatible with the vAxes option.
 * - `'top'` - Above the chart.
 */

/**
 * @typedef {Object} Series An array of objects, each describing the format of the corresponding series in the chart. To use default values for a series, specify an empty object `{}`. If a series or a value is not specified, the global value will be used. 
 * @property {Object} [annotations] An object to be applied to annotations for this series. This can be used to control, for instance, the `textStyle` for the series:
 * ```javascript
 * annotations: {
 *  textStyle: { fontSize: 12, color: 'red' }
 * }
 * ```
 * See the various **annotations** options for a more complete list of what can be customized.
 * @property {string} [color] The color to use for this series. Specify a valid HTML color string.
 * @property {string} [labelInLegend] The description of the series to appear in the chart legend.
 * @property {number} [targetAxisIndex = 0] Which axis to assign this series to, where `0` is the default axis, and `1` is the opposite axis. Default value is `0`; set to `1` to define a chart where different series are rendered against different axes. At least one series must be allocated to the default axis. You can define a different scale for different axes.
 * @property {boolean} [visibleInLegend = true] A boolean value, where `true` means that the series should have a legend entry, and `false` means that it should not. Default is `true`.
 */

/**
 * @typedef {Object} Trendline Displays [trendlines](https://developers.google.com/chart/interactive/docs/gallery/trendlines) on the charts that support them. By default, `linear` trendlines are used, but this can be customized with the **trendlines.*n*.type** option.
 * @property {string} [color] The color of the [trendline](https://developers.google.com/chart/interactive/docs/gallery/trendlines), expressed as either an English color name or a hex string. Defaults to the default series color.
 * @property {number} [degree] For [trendlines](https://developers.google.com/chart/interactive/docs/gallery/trendlines) of `type: 'polynomial'`, the degree of the polynomial (`2` for quadratic, `3` for cubic, and so on).
 * @property {string | null} [labelInLegend = null] If set, the [trendline](https://developers.google.com/chart/interactive/docs/gallery/trendlines) will appear in the legend as this string.
 * @property {number} [lineWidth = 2] The line width of the [trendline](https://developers.google.com/chart/interactive/docs/gallery/trendlines), in pixels.
 * @property {string} [type = 'linear'] Whether the [trendlines](https://developers.google.com/chart/interactive/docs/gallery/trendlines) is 'linear' (the default), 'exponential', or 'polynomial'.
 * @property {boolean} [visibleInLegend = false] Whether the [trendline](https://developers.google.com/chart/interactive/docs/gallery/trendlines) equation appears in the legend. It will appear in the trendline tooltip.
 * 
 * Trendlines are specified on a per-series basis, so most of the time your options will look like this:
 * ```javascript
 * var options = {
 *  trendlines: {
 *   0: {
 *     type: 'linear',
 *     color: 'green',
 *     lineWidth: 3,
 *     opacity: 0.3,
 *     visibleInLegend: true
 *    }
 *  }
 * }
 * ```
 */

/**
 * @typedef {Object} UpdateChartOptions
 * @property {string | {fill: string}} [backgroundColor = 'white'] The background color for the main area of the chart. Can be either a simple HTML color string, for example: `'red'` or `'#00cc00'`, or an object with a `fill` property.
 * @property {ChartArea | null} [chartArea = null] An object with members to configure the placement and size of the chart area (where the chart itself is drawn, excluding axis and legends). Two formats are supported: a number, or a number followed by %. A simple number is a value in pixels; a number followed by % is a percentage. Example: 
 * @property {string[]} [colors] The colors to use for the chart elements. An array of strings, where each element is an HTML color string, for example: `colors:[`'red'`,`'#004411'`]`. Defaults to the default colors.
 * 
 * @property {Axis[] | Object.<number, Axis> | null} [hAxes = null] Specifies properties for individual horizontal axes, if the chart has multiple horizontal axes. Each child object is a `hAxis` object, and can contain all the properties supported by `hAxis`. These property values override any global settings for the same property. \
 * To specify a chart with multiple horizontal axes, first define a new axis using series.targetAxisIndex, then configure the axis using hAxes. The following example assigns series 1 to the bottom axis and specifies a custom title and text style for it:
 * ```javascript 
 * series:{1:{targetAxisIndex:1}}, hAxes:{1:{title:'Losses', textStyle:{color: 'red'}}}
 * ```
 * This property can be either an object or an array: the object is a collection of objects, each with a numeric label that specifies the axis that it defines--this is the format shown above; the array is an array of objects, one per axis. For example, the following array-style notation is identical to the `hAxis` object shown above:
 * ```javascript
 * hAxes: {
 *  {}, // Nothing specified for axis 0
 *  {
 *    title:'Losses',
 *    textStyle: {
 *      color: 'red'
 *    }
 *  } // Axis 1
 * ```
 * 
 * @property {Axis | null} [hAxis = null] An object with members to configure various horizontal axis elements.
 * @property {number} [height = auto] The minimum horizontal data value to render. Ignored when `hAxis.viewWindowMode` is `'pretty'` or `'maximized'`.
 * 
 * @property {boolean | ('percent' | 'relative' | 'absolute')} [isStacked = false] If set to `true`, stacks the elements for all series at each domain value. **Note**: In [Column](https://developers.google.com/chart/interactive/docs/gallery/columnchart), [Area](https://developers.google.com/chart/interactive/docs/gallery/areachart), and [SteppedArea](https://developers.google.com/chart/interactive/docs/gallery/steppedareachart) charts, Google Charts reverses the order of legend items to better correspond with the stacking of the series elements (E.g. series 0 will be the bottom-most legend item). This **does not** apply to [Bar](https://developers.google.com/chart/interactive/docs/gallery/barchart) Charts. \
 * The `isStacked` option also supports 100% stacking, where the stacks of elements at each domain value are rescaled to add up to 100%. \
 * The options for `isStacked` are:
 * - `false` — elements will not stack. This is the default option.
 * - `true` — stacks elements for all series at each domain value.
 * - `'percent'` — stacks elements for all series at each domain value and rescales them such that they add up to 100%, with each element's value calculated as a percentage of 100%.
 * - `'relative'` — stacks elements for all series at each domain value and rescales them such that they add up to 1, with each element's value calculated as a fraction of 1.
 * - `'absolute'` — functions the same as `isStacked: true`. \
 * For 100% stacking, the calculated value for each element will appear in the tooltip after its actual value. \
 * The target axis will default to tick values based on the relative 0-1 scale as fractions of 1 for `'relative'`, and 0-100% for `'percent'` (**Note**: when using the `'percent'` option, the axis/tick values are displayed as percentages, however the actual values are the relative 0-1 scale values. This is because the percentage axis ticks are the result of applying a format of "#.##%" to the relative 0-1 scale values. When using `isStacked: 'percent'`, be sure to specify any ticks/gridlines using the relative 0-1 scale values). You can customize the gridlines/tick values and formatting using the appropriate `hAxis/vAxis` options. \
 * 100% stacking only supports data values of type number, and must have a baseline of zero.
 * 
 * @property {Legend | null} [legend = null] An object with members to configure various aspects of the legend.
 * @property {TextStyle} [legendTextStyle = {color: 'black', fontName: '<global-font-name>', fontSize: '<global-font-size>'}] An object that specifies the legend text style. The `color` can be any HTML color string, for example: `'red'` or '`#00cc00`'. Also see `fontName` and `fontSize`.
 * @property {boolean} [reverseCategories = false] If set to `true`, draws series from right to left. The default is to draw left to right. This option is only supported for a [discrete major](https://developers.google.com/chart/interactive/docs/customizing_axes#Terminology) axis.
 * 
 * @property {Series[] | Object.<number, Series> | null} [series = {}] An array of objects, each describing the format of the corresponding series in the chart. To use default values for a series, specify an empty object `{}`. If a series or a value is not specified, the global value will be used.
 * 
 * You can specify either an array of objects, each of which applies to the series in the order given, or you can specify an object where each child has a numeric key indicating which series it applies to. For example, the following two declarations are identical, and declare the first series as black and absent from the legend, and the fourth as red and absent from the legend:
 * ```javascript
 * series: [
 *    {color: 'black', visibleInLegend: false}, {}, {},
 *    {color: 'red', visibleInLegend: false}
 * ]
 * series: {
 *    0:{color: 'black', visibleInLegend: false},
 *    3:{color: 'red', visibleInLegend: false}
 * }
 * ```
 * 
 * @property {string | null} [subtitle = null] Text to display below the chart title. Set to `undefined` to leave as is, and `null` to remove.
 * @property {TextStyle} [subtitleTextStyle = {color: 'black', fontName: '<global-font-name>', fontSize: '<global-font-size>'}] An object that specifies the title text style. The `color` can be any HTML color string, for example: `'red'` or '`#00cc00`'. Also see `fontName` and `fontSize`.
 * @property {'maximized' | string | null} [theme = null] A theme is a set of predefined option values that work together to achieve a specific chart behavior or visual effect. Currently only one theme is available:
 * - `'maximized'` - Maximizes the area of the chart, and draws the legend and all of the labels inside the chart area.
 * @property {string | null} [title = null] Text to display above the chart.
 * @property {TextStyle} [titleTextStyle = {color: 'black', fontName: '<global-font-name>', fontSize: '<global-font-size>'}] An object that specifies the title text style. The `color `can be any HTML color string, for example: `'red'` or '`#00cc00`'. Also see `fontName` and `fontSize`.
 * @property {Object.<number, Trendline> | null} [trendlines = null] Displays trendlines on the charts that support them. By default, `linear` trendlines are used, but this can be customized with the **trendlines.*n*.type** option.
 * @property {boolean} [useFirstColumnAsDomain] If set to `true`, the chart will treat the column as the domain.
 * @property {Axis | null} [vAxis = null] An object with members to configure various vertical axis elements.
 * @see [Chart Configuartion Options](https://developers.google.com/apps-script/chart-configuration-options)
 */