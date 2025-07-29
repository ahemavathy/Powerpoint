using PowerPointGenerator.Models;
using PowerPointGenerator.Services;

namespace PowerPointGenerator.Examples
{
    /// <summary>
    /// Advanced examples demonstrating various use cases for AI-generated PowerPoint presentations
    /// </summary>
    public static class AdvancedExamples
    {
        /// <summary>
        /// Creates a business report presentation with financial charts and analysis
        /// </summary>
        public static async Task<string> CreateBusinessReportPresentation()
        {
            var content = new PresentationContent
            {
                Title = "Q4 Business Performance Report",
                Author = "AI Business Analyst",
                Slides = new List<SlideContent>()
            };

            // Executive Summary
            content.Slides.Add(new SlideContent
            {
                Title = "Executive Summary",
                Synopsis = "Q4 performance exceeded expectations with 23% revenue growth and improved operational efficiency across all divisions.",
                LayoutType = SlideLayoutType.TitleAndContent,
                BulletPoints = new List<string>
                {
                    "Revenue increased by 23% year-over-year",
                    "Operational costs reduced by 15%",
                    "Customer satisfaction improved to 94%",
                    "Market share expanded in key segments"
                }
            });

            // Financial Performance with Charts
            content.Slides.Add(new SlideContent
            {
                Title = "Financial Performance Overview",
                Synopsis = "Key financial metrics demonstrate strong growth trajectory and improved profitability margins.",
                LayoutType = SlideLayoutType.ImageGrid,
                Images = CreateFinancialChartImages(),
                BulletPoints = new List<string>
                {
                    "Revenue growth acceleration",
                    "Margin improvement",
                    "Cost optimization success"
                }
            });

            // Market Analysis Comparison
            content.Slides.Add(new SlideContent
            {
                Title = "Market Position Analysis",
                Synopsis = "Comparative analysis shows significant improvement in competitive positioning versus last year.",
                LayoutType = SlideLayoutType.TwoImageComparison,
                Images = CreateMarketComparisonImages()
            });

            return await GeneratePresentation(content, "Business_Report_Q4.pptx");
        }

        /// <summary>
        /// Creates a research presentation with scientific data visualization
        /// </summary>
        public static async Task<string> CreateResearchPresentation()
        {
            var content = new PresentationContent
            {
                Title = "Climate Change Impact Analysis",
                Author = "AI Research Assistant",
                Slides = new List<SlideContent>()
            };

            // Research Overview
            content.Slides.Add(new SlideContent
            {
                Title = "Research Methodology",
                Synopsis = "Comprehensive analysis using satellite data from 2010-2024 to assess climate pattern changes.",
                LayoutType = SlideLayoutType.SingleImageWithCaption,
                Images = new List<ImageContent>
                {
                    new ImageContent
                    {
                        FilePath = GetSampleImagePath("methodology_diagram.png"),
                        AltText = "Research methodology flowchart",
                        Caption = "Multi-layered analysis approach combining satellite imagery, ground sensors, and AI pattern recognition to identify climate trends over 14-year period."
                    }
                }
            });

            // Data Visualization
            content.Slides.Add(new SlideContent
            {
                Title = "Temperature Trend Analysis",
                Synopsis = "Regional temperature variations show accelerating change patterns in polar and equatorial regions.",
                LayoutType = SlideLayoutType.ImageFocused,
                Images = CreateResearchDataImages()
            });

            // Results Comparison
            content.Slides.Add(new SlideContent
            {
                Title = "Before vs After Analysis",
                Synopsis = "Comparative visualization of ecosystem changes demonstrates significant environmental impact.",
                LayoutType = SlideLayoutType.TwoImageComparison,
                Images = CreateBeforeAfterImages("ecosystem")
            });

            return await GeneratePresentation(content, "Climate_Research_Analysis.pptx");
        }

        /// <summary>
        /// Creates a product showcase presentation with marketing visuals
        /// </summary>
        public static async Task<string> CreateProductShowcasePresentation()
        {
            var content = new PresentationContent
            {
                Title = "Product Innovation Showcase 2024",
                Author = "AI Product Manager",
                Slides = new List<SlideContent>()
            };

            // Product Overview
            content.Slides.Add(new SlideContent
            {
                Title = "Introducing Next-Gen Solutions",
                Synopsis = "Revolutionary products designed to transform user experience and market expectations.",
                LayoutType = SlideLayoutType.ImageFocused,
                Images = CreateProductHeroImages()
            });

            // Feature Comparison
            content.Slides.Add(new SlideContent
            {
                Title = "Feature Evolution",
                Synopsis = "Side-by-side comparison highlighting significant improvements in functionality and design.",
                LayoutType = SlideLayoutType.TwoImageComparison,
                Images = CreateBeforeAfterImages("product_features")
            });

            // Product Gallery
            content.Slides.Add(new SlideContent
            {
                Title = "Product Ecosystem",
                Synopsis = "Complete product line designed for seamless integration and enhanced user experience.",
                LayoutType = SlideLayoutType.ImageGrid,
                Images = CreateProductGalleryImages(),
                BulletPoints = new List<string>
                {
                    "Seamless cross-platform integration",
                    "Enhanced user interface design",
                    "Advanced performance optimization",
                    "Sustainable manufacturing process"
                }
            });

            return await GeneratePresentation(content, "Product_Showcase_2024.pptx");
        }

        /// <summary>
        /// Creates an educational presentation with instructional diagrams
        /// </summary>
        public static async Task<string> CreateEducationalPresentation()
        {
            var content = new PresentationContent
            {
                Title = "Advanced Machine Learning Concepts",
                Author = "AI Education Assistant",
                Slides = new List<SlideContent>()
            };

            // Concept Introduction
            content.Slides.Add(new SlideContent
            {
                Title = "Neural Network Architecture",
                Synopsis = "Understanding the fundamental building blocks of modern artificial intelligence systems.",
                LayoutType = SlideLayoutType.SingleImageWithCaption,
                Images = new List<ImageContent>
                {
                    new ImageContent
                    {
                        FilePath = GetSampleImagePath("neural_network_diagram.png"),
                        AltText = "Neural network architecture diagram",
                        Caption = "Multi-layer neural network showing input layer, hidden layers, and output layer with weighted connections and activation functions."
                    }
                }
            });

            // Process Flow
            content.Slides.Add(new SlideContent
            {
                Title = "Learning Process Visualization",
                Synopsis = "Step-by-step breakdown of how machine learning models process information and improve over time.",
                LayoutType = SlideLayoutType.ImageGrid,
                Images = CreateEducationalDiagrams(),
                BulletPoints = new List<string>
                {
                    "Data preprocessing and normalization",
                    "Feature extraction and selection",
                    "Model training and validation",
                    "Performance evaluation and optimization"
                }
            });

            return await GeneratePresentation(content, "ML_Educational_Content.pptx");
        }

        /// <summary>
        /// Helper method to generate presentations
        /// </summary>
        private static async Task<string> GeneratePresentation(PresentationContent content, string fileName)
        {
            var outputPath = Path.Combine(Environment.CurrentDirectory, fileName);
            
            using var generator = new PowerPointGeneratorService();
            await generator.CreatePresentationAsync(content, outputPath);
            
            return outputPath;
        }

        /// <summary>
        /// Creates sample financial chart images
        /// </summary>
        private static List<ImageContent> CreateFinancialChartImages()
        {
            return new List<ImageContent>
            {
                new ImageContent
                {
                    FilePath = GetSampleImagePath("revenue_chart.png"),
                    AltText = "Revenue growth chart",
                    Caption = "Quarterly revenue progression"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("profit_margin_chart.png"),
                    AltText = "Profit margin analysis",
                    Caption = "Profit margin improvement over time"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("market_share_chart.png"),
                    AltText = "Market share visualization",
                    Caption = "Market share expansion by segment"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("cost_reduction_chart.png"),
                    AltText = "Cost reduction analysis",
                    Caption = "Operational cost optimization results"
                }
            };
        }

        /// <summary>
        /// Creates market comparison images
        /// </summary>
        private static List<ImageContent> CreateMarketComparisonImages()
        {
            return new List<ImageContent>
            {
                new ImageContent
                {
                    FilePath = GetSampleImagePath("market_position_before.png"),
                    AltText = "Market position 2023",
                    Caption = "Market position Q4 2023"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("market_position_after.png"),
                    AltText = "Market position 2024",
                    Caption = "Market position Q4 2024"
                }
            };
        }

        /// <summary>
        /// Creates research data visualization images
        /// </summary>
        private static List<ImageContent> CreateResearchDataImages()
        {
            return new List<ImageContent>
            {
                new ImageContent
                {
                    FilePath = GetSampleImagePath("temperature_trends.png"),
                    AltText = "Global temperature trend visualization",
                    Caption = "Global temperature anomalies 2010-2024"
                }
            };
        }

        /// <summary>
        /// Creates before/after comparison images for various contexts
        /// </summary>
        private static List<ImageContent> CreateBeforeAfterImages(string context)
        {
            return new List<ImageContent>
            {
                new ImageContent
                {
                    FilePath = GetSampleImagePath($"{context}_before.png"),
                    AltText = $"{context} before state",
                    Caption = "Before implementation"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath($"{context}_after.png"),
                    AltText = $"{context} after state",
                    Caption = "After implementation"
                }
            };
        }

        /// <summary>
        /// Creates product hero images
        /// </summary>
        private static List<ImageContent> CreateProductHeroImages()
        {
            return new List<ImageContent>
            {
                new ImageContent
                {
                    FilePath = GetSampleImagePath("product_hero.png"),
                    AltText = "Product showcase hero image",
                    Caption = "Next-generation product design"
                }
            };
        }

        /// <summary>
        /// Creates product gallery images
        /// </summary>
        private static List<ImageContent> CreateProductGalleryImages()
        {
            return new List<ImageContent>
            {
                new ImageContent
                {
                    FilePath = GetSampleImagePath("product_1.png"),
                    AltText = "Product variant 1",
                    Caption = "Core model"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("product_2.png"),
                    AltText = "Product variant 2",
                    Caption = "Pro model"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("product_3.png"),
                    AltText = "Product variant 3",
                    Caption = "Enterprise model"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("product_4.png"),
                    AltText = "Product ecosystem",
                    Caption = "Complete ecosystem"
                }
            };
        }

        /// <summary>
        /// Creates educational diagram images
        /// </summary>
        private static List<ImageContent> CreateEducationalDiagrams()
        {
            return new List<ImageContent>
            {
                new ImageContent
                {
                    FilePath = GetSampleImagePath("data_preprocessing.png"),
                    AltText = "Data preprocessing diagram",
                    Caption = "Data preparation pipeline"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("feature_extraction.png"),
                    AltText = "Feature extraction process",
                    Caption = "Feature engineering workflow"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("model_training.png"),
                    AltText = "Model training visualization",
                    Caption = "Training optimization process"
                },
                new ImageContent
                {
                    FilePath = GetSampleImagePath("performance_evaluation.png"),
                    AltText = "Performance metrics",
                    Caption = "Model evaluation metrics"
                }
            };
        }

        /// <summary>
        /// Gets sample image path (placeholder for actual images)
        /// </summary>
        private static string GetSampleImagePath(string imageName)
        {
            var imagesDir = Path.Combine(Environment.CurrentDirectory, "Images");
            var imagePath = Path.Combine(imagesDir, imageName);
            
            // In a real implementation, ensure images exist or use actual paths
            return imagePath;
        }
    }
}
