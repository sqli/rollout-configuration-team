import React, { createContext, useState, useContext, useEffect } from "react";

const MarketContext = createContext();

const fmData = {
  "Argentina": {
    B2B: {
      "Store ID": "NesStore_AR_AR_B2B",
      "Shipping Countries": "ar",
      "Web Call Back": 'TRUE',
    },
    B2C: {
      "Store ID": "NesStore_AR_AR",
      "Shipping Countries": "ar",
      "Web Call Back": 'FALSE',
    },
  },
  "Australia": {
    B2B: {
      "Store ID": "NesStore_AU_AU_B2B",
      "Shipping Countries": "au",
      "Web Call Back": 'TRUE',
    },
    B2C: {
      "Store ID": "NesStore_AU_AU",
      "Shipping Countries": "au",
      "Web Call Back": 'FALSE',
    }
  },
  "Austria": {
    B2B: {
      "Store ID": "NesStore_AT_AT_B2B",
      "Shipping Countries": "at",
      "Web Call Back": 'TRUE',
    },
    B2C: {
      "Store ID": "NesStore_AT_AT",
      "Shipping Countries": "at",
      "Web Call Back": 'FALSE',
    }
  }
};

export function MarketProvider({ children }) {
  const [selectedMarkets, setSelectedMarkets] = useState([]);
  const [selectedFeatures, setSelectedFeatures] = useState([]);
  const allMarkets = getAllMarkets();
  const allFeatures = getAllFeatures();

  function getAllMarkets() {
    const allMarkets = [];
    for (const market in fmData) {
      for (const category in fmData[market]) {
        allMarkets.push(`${market}_${category}`);
      }
    }
    return allMarkets;
  }

  function getAllFeatures() {
    const allFeatures = [];
    for (const market in fmData) {
      for (const category in fmData[market]) {
        for (const feature in fmData[market][category]) {
          if (!allFeatures.includes(feature)) {
            allFeatures.push(feature);
          }
        }
      }
    }
    return allFeatures;
  }
  return (
    <MarketContext.Provider
      value={{
        fmData,
        selectedMarkets,
        setSelectedMarkets,
        selectedFeatures,
        setSelectedFeatures,
        allMarkets,
        allFeatures,
      }}
    >
      {children}
    </MarketContext.Provider>
  );
}

export function useMarketContext() {
  return useContext(MarketContext);
}
