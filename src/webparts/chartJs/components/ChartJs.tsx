import {
  ChartControl,
  ChartType,
} from "@pnp/spfx-controls-react/lib/ChartControl";
import * as React from "react";
import { useEffect, useState } from "react";
import { IChartJsProps } from "./IChartJsProps";
import { sp } from "@pnp/sp";

const ChartJs: React.FC<IChartJsProps> = () => {
  const [chartData, setChartData] = useState<any>({});

  const fetchDataFromList = async () => {
    try {
      const response = await sp.web.lists.getByTitle("SalesReport").items.get();
     
      const labels = response.map((item: any) => item.Title);
      const salesData = response.map((item: any) => item.Sale);
      const updatedChartData = {
        labels,
        datasets: [
          {
            label: "Sales Report",
            data: salesData,
          },
        ],
      };

      setChartData(updatedChartData);
    } catch (error) {
      console.error("Error fetching data from SharePoint list:", error);
    }
  };

  useEffect(() => {
    fetchDataFromList();
  }, []); 

  const options = {
    legend: {
      display: true,
      position: "left",
    },
    title: {
      display: true,
      text: "Sales Report Chart",
    },
  };

 if (Object.keys(chartData).length <= 0) {
  <p>Loading...</p>
 }
  return (
    <div>
      {Object.keys(chartData).length > 0 && (
        <ChartControl type={ChartType.Pie} data={chartData} options={options} />
      )}
    </div>
  );
};

export default ChartJs;


// for deployment

// 01. gulp bundle --ship
// 02. gulp package-solution --ship