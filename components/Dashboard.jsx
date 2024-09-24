
import MarketTable from "./MarketTable";
import FilterBar from "./FilterBar";
import Header from "./Header";
import { MarketProvider } from "@/contexts/MarketContext";

export default function Dashboard() {
  return (
    <MarketProvider>
      <div className="container mx-auto p-4">
         <Header/>
        <FilterBar />
        <MarketTable />
      </div>
    </MarketProvider>
  );
}
