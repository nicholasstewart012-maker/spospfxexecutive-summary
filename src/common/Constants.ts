export const AppConstants = {
    Colors: {
        Retail: "#0E2D6D",
        Compliance: "#F9A608",
        Leadership: "#EA5A00",
        Webinars: "#96D4E8",
        Commercial: "#0073B1",
        Other: "#EC008C",
        Default: "#0078D4"
    },
    Categories: {
        Retail: "Retail",
        Compliance: "Compliance Courses",
        Leadership: "Leadership",
        Webinars: "Customer & Emp Oport. Webinars",
        Commercial: "Commercial & Credit",
        Other: "Other"
    }
};

export const getCategoryColor = (category: string): string => {
    switch (category) {
        case AppConstants.Categories.Retail:
            return AppConstants.Colors.Retail;
        case AppConstants.Categories.Compliance:
            return AppConstants.Colors.Compliance;
        case AppConstants.Categories.Leadership:
            return AppConstants.Colors.Leadership;
        case AppConstants.Categories.Webinars:
            return AppConstants.Colors.Webinars;
        case AppConstants.Categories.Commercial:
            return AppConstants.Colors.Commercial;
        case AppConstants.Categories.Other:
            return AppConstants.Colors.Other;
        default:
            return AppConstants.Colors.Other;
    }
};
