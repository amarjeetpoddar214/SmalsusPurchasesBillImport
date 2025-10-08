
import React from 'react';
import Button from './ui/Button';
import { WebPartContext } from '@microsoft/sp-webpart-base';


interface HeaderProps {
  context: WebPartContext;
  selectedMonth: string;
  onPreviousMonth: () => void;
  onNextMonth: () => void;
  onAddPurchase: () => void;
  onAddTransaction: () => void;
  onImportTransactions: () => void;
  onExportData: () => void;
}

const Header: React.FC<HeaderProps> = ({
  context,
  selectedMonth,
  onPreviousMonth,
  onNextMonth,
  onAddPurchase,
  onAddTransaction,
  onImportTransactions,
  onExportData,
}) => {
  const formattedMonth = new Date(selectedMonth + '-02').toLocaleString('default', { month: 'long', year: 'numeric' });
  const dynamicName = context.pageContext.site.absoluteUrl.split('.')[0].replace('https://', '');
  const logo = context.pageContext.web.logoUrl
  console.log(logo, "logoooo")

  return (
    <header className="bg-slate-900/70 backdrop-blur-md sticky top-0 z-50 border-b border-slate-700">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-20 flex-wrap">
          <div className="flex items-center space-x-2">
            {/* <img
              src={`${context.pageContext.site.absoluteUrl}/_api/GroupService/GetGroupImage?id='22e6cdcf-41da-4a5c-ae82-508c5dfeb60d'&hash=637932001185380118`}
              alt="Site Group Image"
              className="h-8 w-8 rounded-full object-cover"
            /> */}
            <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAfCAMAAACxiD++AAAAllBMVEVHcEwrSnbw+Px+zPVqyPWyzuBIX4Msr/BsxfRleJUzsfFEW4FOvPUHqO8GMmU8tPFwhaFtgp9DWoBAtfFHt/E3UXouSnQ5t/EhQ3E1s/AtsPBhv/KGqL9gc5EAp+4AMGQALmIAaMcAqvAAassAY8IAKlwVqu8GiNkAZsYAmeUAPn4cPGsAc84jQm8BUJwONGYAWq0AXLHryHn4AAAAHnRSTlMAlAs7Mh14/Vdi5td1YsDaPzO9yJG/+mPBpZiVZZh2fWC9AAABIUlEQVQokZXQ6VaDMBAF4IDsi6dQu7hlBkKgUGj1/V/OEIIghKPef5zvksmEkH/EtrxTHEdxfPIsa4neS1DSecrS3T9/85NLddkrtvRMS3PwBz1TGqkDNv6nwTh/w0vrlwNex+VXksuMA4i3ZFq1jLXV21ahKi6XLGMND23NknnNhAq/ccBuJwvmvFAXg39yAMB7IhvB9H9VZDJF7yJHWYjWDjgU8H3+EHm9dHGN4ZrlhovGcE9PNuT1MvbBJwd01Kouzdu+wLKGwyxojM8VVRljrLghgr7gc2ga4D95GkHMKwLikkUhUYVwbdKP4wC9A/rTAK0f1AGp1hFD5bu7TrEbNyDhtf9GtQX2+0B38MksduI7hpGeH0XOqeH4Nvl7vgDxS1O9aYSxlQAAAABJRU5ErkJggg==" alt="Site Logo" />
            <h1 className="text-2xl font-bold text-white">{dynamicName}</h1>
          </div>

          <div className="flex items-center gap-2">
            <Button onClick={onImportTransactions} variant="secondary" className="hidden sm:inline-flex">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
              </svg>
              Import
            </Button>
            <Button onClick={onExportData} variant="secondary" className="hidden sm:inline-flex">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
              </svg>
              Export
            </Button>
            <Button onClick={onAddTransaction} variant="secondary">
              Add Transaction
            </Button>
            <Button onClick={onAddPurchase} variant="primary">
              Add Purchase
            </Button>
          </div>
        </div>
        <div className="flex items-center justify-center py-2">
          <Button onClick={onPreviousMonth} size="icon" variant="ghost" aria-label="Previous month">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M15 19l-7-7 7-7" />
            </svg>
          </Button>
          <span className="font-semibold text-lg text-white mx-4 w-40 text-center">{formattedMonth}</span>
          <Button onClick={onNextMonth} size="icon" variant="ghost" aria-label="Next month">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 5l7 7-7 7" />
            </svg>
          </Button>
        </div>
      </div>
    </header>
  );
};

export default Header;
