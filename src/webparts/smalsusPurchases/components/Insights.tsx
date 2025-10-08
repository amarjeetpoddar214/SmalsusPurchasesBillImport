
import React, { useState, useCallback } from 'react';
import { getSpendingInsights } from '../../../services/geminiService';
import { Purchase } from './types';
import Card from './ui/Card';
import Button from './ui/Button';

interface InsightsProps {
  purchases: Purchase[];
}

const Insights: React.FC<InsightsProps> = ({ purchases }) => {
  const [insights, setInsights] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  const handleGenerateInsights = useCallback(async () => {
    setIsLoading(true);
    setInsights('');
    const result = await getSpendingInsights(purchases);
    setInsights(result);
    setIsLoading(false);
  }, [purchases]);

  return (
    <Card>
      <h2 className="text-xl font-bold text-white mb-2">Spending Insights</h2>
      <p className="text-sm text-slate-400 mb-4">Get AI-powered feedback on your spending habits.</p>
      
      {insights && (
        <div className="bg-slate-800/50 p-3 rounded-md mb-4 border border-slate-700">
          <p className="text-slate-200 whitespace-pre-wrap">{insights}</p>
        </div>
      )}

      <Button onClick={handleGenerateInsights} disabled={isLoading || purchases.length < 3} variant="primary" className="w-full">
        {isLoading ? 'Generating...' : 'Generate Insights'}
      </Button>
      {purchases.length < 3 && <p className="text-xs text-slate-500 mt-2 text-center">Add at least 3 purchases to enable insights.</p>}
    </Card>
  );
};

export default Insights;
