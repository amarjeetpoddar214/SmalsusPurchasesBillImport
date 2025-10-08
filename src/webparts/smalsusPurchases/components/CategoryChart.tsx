import React, { useMemo } from 'react';
import { PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { Purchase, Category } from './types';
import Card from './ui/Card';

// Type assertion for Recharts components to fix TypeScript compatibility issues
const ResponsiveContainerFixed = ResponsiveContainer as any;
const PieChartFixed = PieChart as any;
const CellFixed = Cell as any;
const LegendFixed = Legend as any;

interface CategoryChartProps {
  purchases: Purchase[];
}

const COLORS: Record<Category, string> = {
  [Category.Groceries]: '#22c55e', // green-500
  [Category.Utilities]: '#3b82f6', // blue-500
  [Category.Entertainment]: '#a855f7', // purple-500
  [Category.Transport]: '#eab308', // yellow-500
  [Category.Health]: '#ef4444', // red-500
  [Category.Shopping]: '#ec4899', // pink-500
  [Category.FoodAndDining]: '#f97316', // orange-500
  [Category.Other]: '#64748b', // slate-500
};

const CategoryChart: React.FC<CategoryChartProps> = ({ purchases }) => {
  const chartData:any = useMemo(() => {
    const categoryTotals = purchases.reduce((acc, purchase) => {
      // FIX: Ensure `purchase.amount` is treated as a number to prevent string concatenation.
      // This prevents a downstream error in `.sort()` which expects numeric values.
      acc[purchase.category] = (acc[purchase.category] || 0) + Number(purchase.amount);
      return acc;
    }, {} as Record<string, number>);

    return Object.entries(categoryTotals)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
  }, [purchases]);

  if (purchases.length === 0) {
    return (
      <Card>
        <h2 className="text-xl font-bold text-white mb-2">Category Breakdown</h2>
        <div className="flex items-center justify-center h-48">
          <p className="text-slate-400">No data to display chart.</p>
        </div>
      </Card>
    );
  }

  return (
    <Card>
      <h2 className="text-xl font-bold text-white mb-4">Category Breakdown</h2>
      <div style={{ width: '100%', height: 250 }}>
        <ResponsiveContainerFixed>
          <PieChartFixed>
            <Pie
              data={chartData}
              cx="50%"
              cy="50%"
              labelLine={false}
              outerRadius={80}
              fill="#8884d8"
              dataKey="value"
              nameKey="name"
            >
              {chartData.map((entry:any) => (
                <CellFixed key={`cell-${entry.name}`} fill={COLORS[entry.name as Category]} />
              ))}
            </Pie>
            <Tooltip
              contentStyle={{
                backgroundColor: '#1e293b', // slate-800
                borderColor: '#334155', // slate-700
                color: '#f1f5f9', // slate-100
              }}
              formatter={(value: number) => `₹${value.toFixed(2)}`}
            />
            <LegendFixed iconSize={10} />
          </PieChartFixed>
        </ResponsiveContainerFixed>
      </div>
    </Card>
  );
};

export default CategoryChart;
