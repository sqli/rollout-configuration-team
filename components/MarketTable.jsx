import { useMarketContext } from "@/contexts/MarketContext";

export default function MarketTable() {
  const { allMarkets, selectedMarkets, allFeatures, selectedFeatures, fmData } =
    useMarketContext();

  const rowColors = ["#dbdbdb", "#bfbfbf"];
  const colColors = ["#dbffd1", "#c1feae"];

  function renderMarkets() {
    const marketsToRender =
      selectedMarkets.length > 0 ? selectedMarkets : allMarkets;
    return marketsToRender.map((market, index) => (
      <th
        key={index}
        className="p-2 text-gray-700 block md:table-cell"
        style={{
          backgroundColor: colColors[index % colColors.length],
          width: '144px', // Fixed width for column headers
          minWidth: '144px' // Ensure minimum width
        }}
      >
        {market}
      </th>
    ));
  }

  function renderFeatures() {
    const featuresToRender =
      selectedFeatures.length > 0 ? selectedFeatures : allFeatures;
    const marketsToRender =
      selectedMarkets.length > 0 ? selectedMarkets : allMarkets;

    return featuresToRender.map((feature, rowIndex) => (
      <tr
        key={rowIndex}
        className="border border-gray-200 md:table-row"
        style={{ backgroundColor: rowColors[rowIndex % rowColors.length] }} // Apply alternating row colors
      >
        <td
          className="p-2 text-gray-700 block md:table-cell"
          style={{
            width: '144px', // Fixed width for row labels
            minWidth: '144px' // Ensure minimum width
          }}
        >
          {feature}
        </td>
        {marketsToRender.map((m, cellIndex) => {
          const [market, category] = m.split("_");
          const value = fmData[market]?.[category]?.[feature] || "N/A"; // Fallback to 'N/A' if value is not found
          return (
            <td
              key={cellIndex}
              className="p-2 text-gray-700 block md:table-cell text-center"
              style={{
                backgroundColor: colColors[cellIndex % colColors.length],
                width: '144px', // Fixed width for data cells
                minWidth: '144px' // Ensure minimum width
              }}
            >
              {value}
            </td>
          );
        })}
      </tr>
    ));
  }

  return (
    <div className="overflow-auto" style={{ maxHeight: '800px' }}> {/* Adjust height as needed */}
      <table className="border-collapse block md:table table-fixed" style={{ tableLayout: 'fixed' }}>
        <thead className="block md:table-header-group">
          <tr className="border border-gray-200 md:table-row absolute -top-full md:top-auto -left-full md:left-auto md:relative">
            <th className="bg-white p-2 block md:table-cell"
                style={{
                  width: '144px', // Fixed width for header cells
                  minWidth: '144px' // Ensure minimum width
                }}
            ></th>
            {renderMarkets()}
          </tr>
        </thead>
        <tbody className="block md:table-row-group">{renderFeatures()}</tbody>
      </table>
    </div>
  );
}
