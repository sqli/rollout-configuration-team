import { useState } from "react";
import { Select, SelectSection, SelectItem } from "@nextui-org/select";
import { Input } from "@nextui-org/input";
import { Button } from "@nextui-org/react";
import { useMarketContext } from "@/contexts/MarketContext";

export default function FilterBar() {
  const [searchedText,setSearchedText]=useState("");
  const {
    allMarkets,
    selectedMarkets,
    setSelectedMarkets,
    allFeatures,
    selectedFeatures,
    setSelectedFeatures,
  } = useMarketContext();

  function handleSelectMarkets(e) {
    const { value } = e.target;
    const newSelectedMarkets = value ? value.split(",") : [];
    setSelectedMarkets(newSelectedMarkets);
  }

  function handleSelectFeatures(e) {
    const { value } = e.target;
    const newSelectedFeatures = value ? value.split(",") : [];
    setSelectedFeatures(newSelectedFeatures);
  }
  
  function handleSearch(keyword){
    // setSearchedText(keyword);
    // const filteredFeatures=selectedFeatures.filter((feature) =>
    //   feature.toLowerCase().includes(searchedText.toLowerCase())
    // );
    // setSelectedFeatures(filteredFeatures);
  }
  return (
    <div className="flex flex-col md:flex-row items-center gap-4 mb-12">
      <Select
        label="Select a market"
        className="w-80"
        selectionMode="multiple"
        size="sm"
        selectedKeys={selectedMarkets}
        onChange={handleSelectMarkets}
      >
        {allMarkets.map((market) => (
          <SelectItem key={market}>{market}</SelectItem>
        ))}
      </Select>
      <Select
        label="Select a feature"
        className="w-80"
        selectionMode="multiple"
        size="sm"
        selectedKeys={selectedFeatures}
        onChange={handleSelectFeatures}
      >
        {allFeatures.map((feature) => (
          <SelectItem key={feature}>{feature}</SelectItem>
        ))}
      </Select>
      <Input
        className="w-80"
        type="text"
        label="Contains Text"
        placeholder="Search .."
        value={searchedText}
        onChange={(e) => handleSearch(e.target.value)}
      />
      <Button className="w-48" color="default">
        Compare with VST
      </Button>
    </div>
  );
}
