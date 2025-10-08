/**
 * Utility to load Tailwind CSS CDN in SPFx environment
 */
export class TailwindLoader {
  private static isLoaded = false;
  private static stylesOverridden = false;

  public static loadTailwind(): void {
    if (this.isLoaded) {
      return;
    }

    // Check if Tailwind is already loaded
    if (document.querySelector('script[src*="tailwindcss"]')) {
      this.isLoaded = true;
      return;
    }

    // Create and append the script tag
    const script = document.createElement('script');
    script.src = 'https://cdn.tailwindcss.com';
    script.async = true;
    script.onload = () => {
      this.isLoaded = true;
      console.log('Tailwind CSS loaded successfully');
      // Apply SharePoint style overrides after Tailwind loads
      this.overrideCanvasComponent();
    };
    script.onerror = () => {
      console.error('Failed to load Tailwind CSS');
    };

    document.head.appendChild(script);
  }

  public static overrideCanvasComponent(): void {
    if (this.stylesOverridden) {
      return;
    }

    // Create style element to override SharePoint styles
    const style = document.createElement('style');
    style.textContent = `
      /* Override SharePoint CanvasComponent grid styles that conflict with Tailwind */
      .CanvasComponent.LCS .grid:after, 
      .CanvasComponent.LCS .grid:before {
        content: none !important;
        display: none !important;
        line-height: initial !important;
      }
      
      /* Only override SharePoint's grid display when it's NOT a Tailwind grid */
      .CanvasComponent.LCS .grid:not([class*="grid-cols"]):not([class*="grid-rows"]):not([class*="gap"]) {
        display: initial !important;
      }
      
      /* Ensure Tailwind grid classes work properly - be more specific */
      [class*="grid-cols"],
      [class*="grid-rows"],
      .grid[class*="gap"],
      .grid.grid-cols-1,
      .grid.grid-cols-2,
      .grid.grid-cols-3,
      .grid.grid-cols-4,
      .grid.grid-cols-5,
      .grid.grid-cols-6,
      .grid.grid-cols-12 {
        display: grid !important;
      }
      
      /* Responsive grid classes */
      @media (min-width: 640px) {
        [class*="sm:grid-cols"] {
          display: grid !important;
        }
      }
      
      @media (min-width: 768px) {
        [class*="md:grid-cols"] {
          display: grid !important;
        }
      }
      
      @media (min-width: 1024px) {
        [class*="lg:grid-cols"] {
          display: grid !important;
        }
      }
      
      @media (min-width: 1280px) {
        [class*="xl:grid-cols"] {
          display: grid !important;
        }
      }
      
      /* Remove pseudo-elements from all grid classes */
      [class*="grid"]:before,
      [class*="grid"]:after {
        content: none !important;
        display: none !important;
      }
    `;
    
    document.head.appendChild(style);
    this.stylesOverridden = true;
    console.log('SharePoint CanvasComponent styles overridden for Tailwind compatibility');
  }

  public static ensureTailwindLoaded(): Promise<void> {
    return new Promise((resolve) => {
      if (this.isLoaded) {
        resolve();
        return;
      }

      this.loadTailwind();
      
      // Check periodically if Tailwind is loaded
      const checkInterval = setInterval(() => {
        if (this.isLoaded || document.querySelector('script[src*="tailwindcss"]')) {
          clearInterval(checkInterval);
          this.isLoaded = true;
          resolve();
        }
      }, 100);

      // Timeout after 5 seconds
      setTimeout(() => {
        clearInterval(checkInterval);
        resolve();
      }, 5000);
    });
  }
}
