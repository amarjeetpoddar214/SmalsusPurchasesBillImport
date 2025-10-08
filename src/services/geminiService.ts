
import { GoogleGenAI } from "@google/genai";
import { Purchase, Category } from "../webparts/smalsusPurchases/components/types";

// For SPFx, we'll use a placeholder API key for now
// In production, you should configure this through SharePoint's property pane or tenant settings
const API_KEY = process.env.API_KEY || 'your-api-key-here';

if (!API_KEY || API_KEY === 'your-api-key-here') {
    console.warn("Gemini API key not configured. AI features will not work.");
}

const ai = new GoogleGenAI({ apiKey: API_KEY });

const validCategories = Object.values(Category);

export const suggestCategory = async (purchaseName: string): Promise<Category | null> => {
    if (!purchaseName.trim()) return null;

    // Check if API key is properly configured
    if (!API_KEY || API_KEY === 'your-api-key-here') {
        console.warn("Gemini API key not configured. Category suggestion disabled.");
        return null;
    }

    try {
        const prompt = `Analyze the item "${purchaseName}". From the following list of categories, which one is the most appropriate? [${validCategories.join(', ')}]. Respond with ONLY the category name and nothing else.`;
        
        const response:any = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: {
                // Disable thinking for low latency
                thinkingConfig: { thinkingBudget: 0 }
            }
        });
        
        const suggested = response.text.trim();

        // Check if the suggested category is a valid enum member
        if (validCategories.includes(suggested as Category)) {
            return suggested as Category;
        }

        return null;
    } catch (error) {
        console.error("Error suggesting category:", error);
        return null;
    }
};

export const getSpendingInsights = async (purchases: Purchase[]): Promise<string> => {
    if (purchases.length < 3) {
        return "Add more purchases to generate spending insights.";
    }

    // Check if API key is properly configured
    if (!API_KEY || API_KEY === 'your-api-key-here') {
        return "AI insights are not available. Please configure the Gemini API key to enable this feature.";
    }

    try {
        const purchaseList = purchases.map(p => `- ${p.name}: ₹${p.amount.toFixed(2)} (Category: ${p.category})`).join('\n');
        
        const prompt = `
            You are a friendly financial advisor. Based on the following purchase data, provide a brief, actionable insight into the user's spending habits. 
            Mention the highest spending category and suggest one specific, practical area for potential savings.
            Keep the response concise (2-3 sentences) and encouraging.

            Purchases:
            ${purchaseList}
        `;

        const response:any = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
        });

        return response.text;
    } catch (error) {
        console.error("Error getting spending insights:", error);
        return "Could not retrieve insights at this time. Please try again later.";
    }
};
