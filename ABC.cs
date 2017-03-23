using NUnit.Framework;
using OpenQA.Selenium;
using RetailChannels.BF.Common;
using RetailChannels.BF.PageObjects;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using static RetailChannels.BF.Common.Enums;
using static RetailChannels.BF.Common.Types;
using static RetailChannels.BF.PageObjects.ProductDetailsPage;

namespace RetailChannels.BF.Tests.Regression
{
    [TestFixtureSource(typeof(Configs), "StandardBrowsers")]
    // [TestFixture(Configs.Browsers.CHROME, "", Configs.Platform.WINDOWS8)]
    [Parallelizable]
    public class ABC
    {
        private bool localRun = false;

        private static Sites site = Sites.americanbankchecks;
        private static Subdomains subDomain = Subdomains.dev; // DEV, QA
        private static Browsers browser = Browsers.Chrome;

        // SauceLabs
        private IWebDriver Driver;
        private string browserName;
        private string version;
        private string platform;

        public ABC(Configs.Browsers browserName, string version, Configs.Platform platform)
        {
            this.browserName = browserName.ToDescriptionString();
            this.version = version;
            this.platform = platform.ToDescriptionString();
        }

        [SetUp]
        public void Setup()
        {
            if (localRun)
            {
                Driver = Functions.StartLocalDriver(Browsers.Chrome);
            }
            else
            {
                Driver = Functions.StartSauceLabsDriver(browserName, version, platform);
            }
        }

        [Test, Property("Description", "Validate Address Label order as a guest checkout.")]
        public void TC8034_Accessory_AddressLabel_Guest()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "103981";
            List<string> messageLines = new List<string>(new string[] { "Line1" });
            Customer customer = Samples.Customers.CONUS;

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - PRODUCT OPTIONS
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            productDetailsPage.ProductOptions.SetMessageLines(messageLines);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            string price = productDetailsPage.ProductOptions.GetPrice();
            productDetailsPage.ProductOptions.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(price, shoppingCartPage.GetProductPrice(0), "Validate product price");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChoicePage = new MyAccountChoicePage(Driver);
            myAccountChoicePage.ContinueAsGuest();

            // BILLING ADDRESS PAGE
            BillingAddressPage billingAddressPage = new BillingAddressPage(Driver);
            billingAddressPage.EnterAddress(customer);
            billingAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            Assert.True(messageLines.SequenceEqual(shippingDetailsPage.GetMessageLines()));
            string shippingPrice = shippingDetailsPage.GetSelectedShippingPrice();
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            Assert.AreEqual(shippingPrice, paymentPage.GetShippingCost(), "Verify shipping cost");
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(shippingPrice, orderConfirmationPage.GetShippingPrice(), "Verify shipping price");
        }
        [Test, Property("Description", "Validate Checkbook Cover Order as a returning customer.")]
        public void TC8035_Accessory_CheckbookCover_ReturningCustomer()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "98573"; // TODO: Use 103999 once DE7429 is fixed
            Customer customer = Samples.Customers.OCONUS2;

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - PRODUCT OPTIONS
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            productDetailsPage.ProductOptions.SetMessageLines(new List<string>() { "carrot" }); // TODO: remove once DE7429 is fixed
            string productPrice = productDetailsPage.ProductOptions.GetPrice();
            productDetailsPage.ProductOptions.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(productPrice, shoppingCartPage.GetProductPrice(0), "Validate product price");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChoicePage = new MyAccountChoicePage(Driver);
            myAccountChoicePage.ContinueAsGuest();

            // BILLING ADDRESS PAGE
            BillingAddressPage billingAddressPage = new BillingAddressPage(Driver);
            billingAddressPage.EnterAddress(customer);
            billingAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            Assert.AreEqual(productPrice, shippingDetailsPage.GetPrice(), "Validate product price");
            string shippingPrice = shippingDetailsPage.GetSelectedShippingPrice();
            Assert.True(shippingDetailsPage.GetSelectedShippingMethodLabel().Contains("Priority"), "Validate that the shipping method is Priority for AK & HI");
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            Assert.AreEqual(shippingPrice, paymentPage.GetShippingCost(), "Verify shipping cost");
            Assert.True(paymentPage.GetShippingMethod(0).Equals("Priority"), "Validate that the shipping method is Priority for AK & HI");
            Assert.AreEqual(productPrice, paymentPage.GetProductPrice(0), "Verify product price");
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(shippingPrice, orderConfirmationPage.GetShippingPrice(), "Verify shipping price");
            Assert.AreEqual(productPrice, orderConfirmationPage.GetProductPrice(0), "Verify product price");
            Assert.True(orderConfirmationPage.GetShippingMethod(0).Equals("Priority"), "Validate that the shipping method is Priority for AK & HI");
        }
        [Test, Property("Description", "Validate Monogram Stamp Order as a non-member and later signing in as a returning customer.")]
        public void TC8039_Accessory_MonogramStamp_NewCustomer()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "98117";
            List<string> messageLines = new List<string>(new string[] { "Line1" });
            Customer customer = Samples.Customers.CONUS;
            customer.emailAddress = Functions.GenerateEmailAddress();
            ProductStyleColors productColor = ProductStyleColors.Blue;

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - PRODUCT OPTIONS
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            string productPrice = productDetailsPage.ProductOptions.GetPrice();
            // TODO: Can't validate monogram image
            productDetailsPage.ProductOptions.SetStyleColor(productColor);
            productDetailsPage.ProductOptions.SetMessageLines(messageLines);
            productDetailsPage.ProductOptions.SetMonogram("D - Modern");
            productDetailsPage.ProductOptions.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(productPrice, shoppingCartPage.GetProductPrice(0), "Validate product price");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChoicePage = new MyAccountChoicePage(Driver);
            myAccountChoicePage.NewCheckCustomer(customer.emailAddress);

            // MY ACCOUNT PAGE
            MyAccountPage myAccountPage = new MyAccountPage(Driver);
            myAccountPage.CreateAccount(customer);

            // BILLING ADDRESS PAGE
            BillingAddressPage billingAddressPage = new BillingAddressPage(Driver);
            billingAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            Assert.AreEqual(productPrice, shippingDetailsPage.GetPrice(), "Validate product price");
            Assert.AreEqual(messageLines, shippingDetailsPage.GetMessageLines(), "Validate message lines");
            string shippingPrice = shippingDetailsPage.GetSelectedShippingPrice();
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            Assert.AreEqual(shippingPrice, paymentPage.GetShippingCost(), "Verify shipping cost");
            Assert.AreEqual(productPrice, paymentPage.GetProductPrice(0), "Verify product price");
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(shippingPrice, orderConfirmationPage.GetShippingPrice(), "Verify shipping price");
            Assert.AreEqual(productPrice, orderConfirmationPage.GetProductPrice(0), "Verify product price");
        }
        [Test, Property("Description", "Validate Search Functionality")]
        public void TC8040_Search()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string searchString = "blue safety";

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // HOME PAGE
            HomePage homePage = new HomePage(Driver);
            homePage.ClosePopup();

            // HEADER PANEL
            HeaderPanel header = new HeaderPanel(Driver);
            header.Search(searchString);

            // SEARCH RESULTS PAGE
            SearchResultsPage searchResulsPage = new SearchResultsPage(Driver);
            Assert.AreNotEqual(0, searchResulsPage.ResultsCount(), "Verify that search results exist");
            Assert.AreEqual(searchString, searchResulsPage.GetSearchString(), "Verify the search string");

            // TODO: Can't really validate the pagination, etc at the bottom of the page because I can't guarantee that there will be that many results.
        }
        [Test, Property("Description", "Validate Address Net web service.")]
        public void TC8216_ValidateAddressNet()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "98573"; // TODO: Use 103999 once DE7429 is fixed
            Customer customer = Samples.Customers.OCONUS2;
            Customer badCustomer = new Customer("testaccount1@citm.com", false, "1234567890", "John OCONUS", "Grand Wailea", true, "abcd", "", "Wailea", "HI", "96753", "Optional phone line");

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - PRODUCT OPTIONS
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            productDetailsPage.ProductOptions.SetMessageLines(new List<string>() { "carrot" }); // TODO: remove once DE7429 is fixed
            productDetailsPage.ProductOptions.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChoicePage = new MyAccountChoicePage(Driver);
            myAccountChoicePage.ContinueAsGuest();

            // BILLING ADDRESS PAGE
            BillingAddressPage billingAddressPage = new BillingAddressPage(Driver);
            billingAddressPage.EnterAddress(badCustomer);
            billingAddressPage.Continue();
            billingAddressPage.PleaseSelectOne(BillingAddressPage.AddressChangeSelections.Change);
            billingAddressPage.EnterAddress(customer);
            billingAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            string shipToAddress = $"{customer.streetAddress}, {customer.optionalSecondLine}, {customer.city}, {customer.state} {customer.zip} United States";
            Assert.AreEqual(shipToAddress, shippingDetailsPage.GetShipToAddress(0), "Validate ship to address");
        }
        [Test, Property("Description", "Validate Order History Functionality")]
        public void TC8220_OrderHistory()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            Login login = new Login("susmitha.kolachina@harlandclarke.com", "Testing1");

            string orderNumber1 = "39-12690279";
            string phoneNumber1 = "8479972056";
            string orderNumber2 = "39-12690300";
            string accountNumber2 = "0122334455";

            Driver.Manage().Window.Maximize();
            string url = BrowseablePages.orderHistoryPage;
            Driver.Url = url;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // ORDER HISTORY PAGE
            OrderHistoryPage orderHistoryPage = new OrderHistoryPage(Driver);
            orderHistoryPage.LookupSingleOrder(orderNumber: orderNumber1, phoneNumber: phoneNumber1);

            // ORDER DETAIL PAGE
            OrderDetailPage orderDetailPage = new OrderDetailPage(Driver);
            Assert.AreNotEqual(0, orderDetailPage.GetOrderHistoryCount(), "Validate order history is not 0");

            // navigate back to the order history page
            Driver.Navigate().GoToUrl(url);

            // ORDER HISTORY PAGE
            orderHistoryPage = new OrderHistoryPage(Driver);
            orderHistoryPage.LookupSingleOrder(orderNumber: orderNumber2, checkingAccountNumber: accountNumber2);

            // ORDER DETAIL PAGE
            orderDetailPage = new OrderDetailPage(Driver);
            Assert.AreNotEqual(0, orderDetailPage.GetOrderHistoryCount(), "Validate order history is not 0");

            // navigate back to the order history page
            Driver.Navigate().GoToUrl(url);

            // ORDER HISTORY PAGE
            orderHistoryPage = new OrderHistoryPage(Driver);
            orderHistoryPage.SignIn(login);

            // MY ACCOUNT DASHBOARD PAGE
            MyAccountDashboardPage myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            myAccountDashboardPage.CheckOrderStatusOrReorderItems();

            // ORDER DETAIL PAGE
            OrderListingPage orderListingPage = new OrderListingPage(Driver);
            orderListingPage.SearchHere(orderNumber: orderNumber1, phoneNumber: phoneNumber1);
            Assert.AreNotEqual(0, orderListingPage.GetOrderHistoryCount(), "Validate order history is not 0");

            // navigate back to the order history page
            Driver.Navigate().GoToUrl(url);

            // ORDER LISTING PAGE
            orderListingPage.SearchHere(orderNumber: orderNumber2, checkingAccountNumber: accountNumber2);
            Assert.AreNotEqual(0, orderListingPage.GetOrderHistoryCount(), "Validate order history is not 0");
        }
        [Test, Property("Description", "Validate Home Page Header Links functionality")]
        public void TC8037_HeaderLinks()
        {
            BrowseablePages.init(site, subDomain);

            Driver.Manage().Window.Maximize();
            string url = BrowseablePages.domain;
            Driver.Url = url;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // many of the validations are instantiating the page object which checks the URL

            // HOME PAGE
            HomePage homePage = new HomePage(Driver);
            homePage.ClosePopup();
            HeaderPanel header = new HeaderPanel(Driver);

            // SIGN IN
            header.SignIn();
            LoginPage loginPage = new LoginPage(Driver);

            // MY ACCOUNT
            header.MyAccount();
            loginPage = new LoginPage(Driver);

            // ORDER STATUS
            header.OrderStatus();
            OrderHistoryPage orderHistoryPage = new OrderHistoryPage(Driver);

            // SHOPPING CART
            header.ShoppingCart();
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);

            // HELP
            header.Help();
            HelpPage helpPage = new HelpPage(Driver);

            // CONTACT US
            header.ContactUs();
            ContactUsPage contactUsPage = new ContactUsPage(Driver);

            // TOP NAV

            // PERSONAL PRODUCTS

            // PERSONAL CHECKS
            string catalog = "Personal Products";
            string category = "Bestsellers";
            header.ClickAbcTopNav(catalog, category);
            CategoryPage categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Characters Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Animals Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Classic";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Contemporary Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Eco-Friendly Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Floral";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Girl Power Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Inspirational Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Miscellaneous";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Patterns";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Scenic Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Top Stub";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "High Security";
            header.ClickAbcTopNav(catalog, category);
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.AreEqual("High Security Checks", productDetailsPage.GetProductName());

            category = "View All";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            // ADDRESS LABELS
            category = "Bestsellers";
            header.ClickAbcTopNav(catalog, category, 1);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Icons/Characters";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Animals";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Classic";
            header.ClickAbcTopNav(catalog, category, 1);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Contemporary";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Eco-Friendly";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Floral";
            header.ClickAbcTopNav(catalog, category, 1);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Girl Power";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Inspirational";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Miscellaneous";
            header.ClickAbcTopNav(catalog, category, 1);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Patterns";
            header.ClickAbcTopNav(catalog, category, 1);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Scenic";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "View All";
            header.ClickAbcTopNav(catalog, category, 1);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            // ACCESSORIES
            category = "Checkbook Covers";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Ink Stamps";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Deposits/Registers";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            // BUSINESS PRODUCTS
            // COMPUTER CHECKS
            catalog = "Business Products";
            category = "High Security Business Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Laser Voucher - Check on Top";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Laser Voucher - Check In Middle";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Three to a Page";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Laser Wallet";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            // MANUAL CHECKS
            category = "High Security";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Home Desk Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Home Desk Registers";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "General Purpose";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "3 Per Page Business Registers";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Payroll & Voucher Checks";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            // BUSINESS ACCESSORIES
            category = "Deposits/Registers";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Binders & Organizers";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Ink Stamps";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Envelopes";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Specialty Products";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());

            category = "Tax Forms";
            header.ClickAbcTopNav(catalog, category);
            categoryPage = new CategoryPage(Driver);
            Assert.AreEqual(category, categoryPage.GetSelectedLeftNavCategory());
        }
        [Test, Property("Description", "Validate Home Page Footer Links functionality.")]
        public void TC8038_FooterLinks()
        {
            BrowseablePages.init(site, subDomain);

            Driver.Manage().Window.Maximize();
            string url = BrowseablePages.shoppingCartPage;
            Driver.Url = url;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // many of the validations are instantiating the page object which checks the URL

            // SUPPORT & SERVICES
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            FooterPanel footerPanel = new FooterPanel(Driver);

            string link = "Order Status";
            footerPanel.ClickLink(link);
            HelpPage helpPage = new HelpPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20318#tabs-8"), $"Verify {link} link");

            link = "Placing an Order";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            helpPage = new HelpPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20318#tabs-5"), $"Verify {link} link");

            link = "Returns";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            helpPage = new HelpPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20318#tabs-7"), $"Verify {link} link");

            link = "Shipping info";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            helpPage = new HelpPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20318#tabs-9"), $"Verify {link} link");

            link = "Secure Shopping";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            helpPage = new HelpPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20318#tabs-6"), $"Verify {link} link");

            link = "Help";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            helpPage = new HelpPage(Driver);

            link = "Customer Service";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            ContactUsPage contactUsPage = new ContactUsPage(Driver);

            link = "Privacy & Security";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            PrivacyPolicyPage privacyPolicyPage = new PrivacyPolicyPage(Driver);

            link = "Contact Us";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            contactUsPage = new ContactUsPage(Driver);

            link = "Sitemap";
            Driver.Url = url; // Help pages don't have the footer, navigate to a page that does
            footerPanel.ClickLink(link);
            SitemapPage sitemapPage = new SitemapPage(Driver);

            // PRODUCT INFO
            link = "Personal Checks";
            footerPanel.ClickLink(link);
            CategoryPage categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20268/"), $"Verify {link} link");

            link = "Deposits & Registers";
            footerPanel.ClickLink(link);
            categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20301/"), $"Verify {link} link");

            link = "Address Labels";
            footerPanel.ClickLink(link);
            categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20284/"), $"Verify {link} link");

            link = "Checkbook Covers";
            footerPanel.ClickLink(link);
            categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20299/"), $"Verify {link} link");

            link = "Custom Ink Stamps";
            footerPanel.ClickLink(link);
            categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20313/"), $"Verify {link} link");

            link = "Envelopes";
            footerPanel.ClickLink(link);
            categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20314/"), $"Verify {link} link");

            link = "Computer Checks";
            footerPanel.ClickLink(link);
            categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20303/"), $"Verify {link} link");

            link = "Home Desk Checks";
            footerPanel.ClickLink(link);
            categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20305/"), $"Verify {link} link");

            link = "Binders & Organizers";
            footerPanel.ClickLink(link);
            categoryPage = new CategoryPage(Driver);
            Assert.IsTrue(Driver.Url.Contains("/20312/"), $"Verify {link} link");
        }
        [Test, Property("Description", "Validate My Account Functionality.")]
        public void TC8217_MyAccount()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            Customer customer = Samples.Customers.CONUS;
            Customer originalCustomer = new Customer("ksussy@gmail.com", false, "8479972056", "Susmitha Kolachina", "Harland Clarke", false, "2435 Goodwin Lane", "", "New Braunfels", "TX", "78135", "");
            Login login = new Login("ksussy@gmail.com", "Testing1");
            string newPassword = "Testing2";
            string originalCompanyName = "Harland Clarke";
            string newCompanyName = "Harland Clarke 2";

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // HOME PAGE
            HomePage homePage = new HomePage(Driver);
            homePage.ClosePopup();
            HeaderPanel header = new HeaderPanel(Driver);
            header.SignIn();

            // LOGIN PAGE
            LoginPage loginPage = new LoginPage(Driver);
            loginPage.SignIn(login);

            // TODO: Step 6 is covered in another test case, remove?

            // STEP - CHANGE PASSWORD

            // MY ACCOUNT DASHBOARD PAGE
            MyAccountDashboardPage myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            myAccountDashboardPage.ChangePassword();

            // CHANGE PASSWORD PAGE
            ChangePasswordPage changePasswordPage = new ChangePasswordPage(Driver);
            changePasswordPage.ChangePassword(login.password, newPassword);

            // MY ACCOUNT DASHBOARD PAGE
            myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            Assert.AreEqual("Your Password has been changed successfully.", myAccountDashboardPage.GetMessage(), "Verify successful password change message.");
            myAccountDashboardPage.ChangePassword();

            // CHANGE PASSWORD PAGE
            changePasswordPage = new ChangePasswordPage(Driver);
            changePasswordPage.ChangePassword(newPassword, login.password);

            // MY ACCOUNT DASHBOARD PAGE
            myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            Assert.AreEqual("Your Password has been changed successfully.", myAccountDashboardPage.GetMessage(), "Verify successful password change message.");

            // STEP - MANAGE SHIPPING ADDRESSES

            myAccountDashboardPage.ManageShippingAddresses();

            // MANAGE SHIPPING ADDRESSES PAGE
            ManageShippingAddressesPage manageShippingAddressesPage = new ManageShippingAddressesPage(Driver);
            manageShippingAddressesPage.AddShippingAddress(customer);

            Assert.AreEqual(2, manageShippingAddressesPage.GetShippingAddressCount(), "Verify shipping address was added");
            Assert.AreEqual(customer.name, manageShippingAddressesPage.GetShipName(1), "Validate added shipping name");
            Assert.AreEqual(customer.streetAddress, manageShippingAddressesPage.GetShipAddressLine1(1), "Validate added shipping address line 1");
            Assert.AreEqual(customer.optionalSecondLine, manageShippingAddressesPage.GetShipAddressLine2(1), "Validate added shipping address line 2");
            Assert.AreEqual($"{customer.city}, {customer.state} {customer.zip}", manageShippingAddressesPage.GetShipCityStateZip(1), "Validate added shipping city, state zip");
            Assert.AreEqual("United States", manageShippingAddressesPage.GetShipCountry(1), "Validate added shipping country");
            Assert.AreEqual(customer.daytimePhoneNumber, Regex.Replace(manageShippingAddressesPage.GetShipPhoneNumber(1), @"[^\d]", ""), "Validate added shipping phone number");
            Assert.AreEqual($"Email: {customer.emailAddress}", manageShippingAddressesPage.GetShipEmail(1), "Validate added shipping email");

            manageShippingAddressesPage.DeleteAddress(1);
            Assert.AreEqual(1, manageShippingAddressesPage.GetShippingAddressCount(), "Verify shipping address was deleted");
            manageShippingAddressesPage.BackToMyAccount();

            // STEP - CHANGE BILLING ADDRESS

            // MY ACCOUNT DASHBOARD PAGE
            myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            myAccountDashboardPage.ChangeBillingAddress();

            // BILLING ADDRESS PAGE
            BillingAddressPage billingAddressPage = new BillingAddressPage(Driver);
            billingAddressPage.EnterAddress(customer);
            billingAddressPage.SetCompanyName(newCompanyName);
            billingAddressPage.Continue();

            // MY ACCOUNT DASHBOARD PAGE
            myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            myAccountDashboardPage.ChangeBillingAddress();

            // BILLING ADDRESS PAGE
            billingAddressPage = new BillingAddressPage(Driver);
            Assert.AreEqual(newCompanyName, billingAddressPage.GetCompanyName(), "Verify company name was changed");
            billingAddressPage.SetCompanyName(originalCompanyName);
            billingAddressPage.Continue();

            // MY ACCOUNT DASHBOARD PAGE
            myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            myAccountDashboardPage.ChangeBillingAddress();

            // BILLING ADDRESS PAGE
            billingAddressPage = new BillingAddressPage(Driver);
            Assert.AreEqual(originalCompanyName, billingAddressPage.GetCompanyName(), "Verify company name was changed back");
            billingAddressPage.Cancel();

            // STEP - EDIT ACCOUNT PROFILE

            // MY ACCOUNT DASHBOARD PAGE
            myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            myAccountDashboardPage.EditAccountProfile();

            // MY ACCOUNT PAGE
            MyAccountPage myAccountPage = new MyAccountPage(Driver);
            myAccountPage.SetCompanyName(newCompanyName);
            myAccountPage.Continue();

            // MY ACCOUNT DASHBOARD PAGE
            myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            myAccountDashboardPage.EditAccountProfile();

            // MY ACCOUNT PAGE
            myAccountPage = new MyAccountPage(Driver);
            Assert.AreEqual(newCompanyName, myAccountPage.GetCompanyName(), "Verify company name was changed");
            myAccountPage.SetCompanyName(originalCompanyName);
            myAccountPage.Continue();

            // MY ACCOUNT DASHBOARD PAGE
            myAccountDashboardPage = new MyAccountDashboardPage(Driver);
            myAccountDashboardPage.EditAccountProfile();

            // MY ACCOUNT PAGE
            myAccountPage = new MyAccountPage(Driver);
            Assert.AreEqual(originalCompanyName, myAccountPage.GetCompanyName(), "Verify company name was changed back");
        }
        [Test, Property("Description", "Validate Personal Checks Order with Fraud Armor and continental shipping")]
        public void TC8041_Personal_Checks_Singles_With_Fraud_Armor()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "98154";
            Account account = new Account("314089681", "0122334455");
            string bottomRowOfNumbers = "31408968101223344551001";
            string startingCheckNumber = "1001";
            Customer customer = Samples.Customers.CONUS;
            BankInfo bankInfo = Samples.Banks.Bank1;

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - CUSTOMIZE
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            productDetailsPage.Customize.SetCheckType(CustomizePanel.CheckTypes.Singles);
            productDetailsPage.Customize.SetQuantity("4 Boxes");
            string productPrice = productDetailsPage.Customize.GetPrice();
            productDetailsPage.Customize.AddLogo(LogoTypes.Stock, 1);
            string logoPrice = productDetailsPage.Customize.GetLogoPrice();
            productDetailsPage.Customize.Continue();

            // PRODUCT DETAILS PAGE - PERSONALIZE
            productDetailsPage.Personalize.SetName(customer.name);
            productDetailsPage.Personalize.SetAddressLine(customer.streetAddress);
            productDetailsPage.Personalize.SetCity(customer.city);
            productDetailsPage.Personalize.SetState(customer.state);
            productDetailsPage.Personalize.SetZip(customer.zip);
            productDetailsPage.Personalize.SetOptionalLine(customer.optionalSecondLine);
            productDetailsPage.Personalize.SetFont(1);
            productDetailsPage.Personalize.SetOslLine1("OSL Line 1");
            string fontPrice = productDetailsPage.Personalize.GetFontPrice();
            string oslPrice = productDetailsPage.Personalize.GetOslPrice();
            productDetailsPage.Personalize.Continue();

            // PRODUCT DETAILS PAGE - ACCOUNT
            productDetailsPage.Account.SetBankName(bankInfo.bankName);
            productDetailsPage.Account.SetStartingCheckNumber(startingCheckNumber);
            productDetailsPage.Account.SetRoutingNumber(account.routingNumber);
            productDetailsPage.Account.SetAccountNumber(account.accountNumber);
            productDetailsPage.Account.SetBottomRowofNumbers(bottomRowOfNumbers);
            productDetailsPage.Account.Next();

            // PRODUCT DETAILS PAGE - FORMAT SELECTION DIALOG
            ProductDetailsPage_FormatSelectionDialog formatSelectionDialog = new ProductDetailsPage_FormatSelectionDialog(Driver);
            formatSelectionDialog.Continue();

            // CONFIRMATION PANEL
            ConfirmationPanel confirmationPanel = new ConfirmationPanel(Driver);
            confirmationPanel.SetFraudArmor(true);
            // TODO: Cannot validate Fraud Armor Protection because the price is changing, e.g. $5 per box on PDP and $15 on cart
            confirmationPanel.ClickVerifyInformationCheckbox();
            confirmationPanel.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(productPrice, shoppingCartPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, shoppingCartPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shoppingCartPage.GetLogoPrice(0), "Validate the Logo price");
            Assert.AreEqual(oslPrice, shoppingCartPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", shoppingCartPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChiocePage = new MyAccountChoicePage(Driver);
            myAccountChiocePage.ContinueAsGuest();

            // BILLING ADDRESS PAGE
            BillingAddressPage billAddressPage = new BillingAddressPage(Driver);
            billAddressPage.EnterAddress(customer);
            billAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            Assert.AreEqual(productPrice, shippingDetailsPage.GetProductPrice(0), "Validate the product selected price");
            Assert.AreEqual(fontPrice, shippingDetailsPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shippingDetailsPage.GetLogoPrice(0), "Validate the Logo price");
            Assert.AreEqual(oslPrice, shippingDetailsPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", shippingDetailsPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
            Assert.AreEqual(customer.name, shippingDetailsPage.GetLine(1), "Verify customer name");
            Assert.AreEqual(customer.streetAddress, shippingDetailsPage.GetLine(3), "Verify address");
            Assert.AreEqual($"{customer.city}, {customer.state} {customer.zip}", shippingDetailsPage.GetLine(4), "Verify city, state, zip");
            Assert.AreEqual(customer.optionalSecondLine, shippingDetailsPage.GetLine(5), "Validate the optional line");
            Assert.AreEqual(bankInfo.bankName, shippingDetailsPage.GetBankName(), "Verify bank name");
            Assert.AreEqual(account.routingNumber, shippingDetailsPage.GetBankRoutingNumber(), "Verify routing number");
            Assert.AreEqual(account.accountNumber, shippingDetailsPage.GetCheckingAccountNumber(), "Verify checking account number");
            Assert.AreEqual(startingCheckNumber, shippingDetailsPage.GetCheckStartingNumber(), "Verify check starting number");
            shippingDetailsPage.SetShippingMethod(ShippingDetailsPage.ShippingMethods.Overnight);
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(productPrice, orderConfirmationPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, orderConfirmationPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, orderConfirmationPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, orderConfirmationPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", orderConfirmationPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
        }
        [Test, Property("Description", "Validate Personal Checks Order - Duplicates without Fraud Armor and Priority Shipping.")]
        public void TC8042_Personal_Checks_Duplicates_Without_FraudArmor_Priority()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "98191";
            Account account = new Account("314089681", "0122334455");
            string bottomRowOfNumbers = "31408968101223344551001";
            string startingCheckNumber = "1001";
            Customer customer = Samples.Customers.OCONUS2;
            Login login = new Login("susmitha.kolachina@harlandclarke.com", "Testing1");
            BankInfo bankInfo = Samples.Banks.Bank1;
            string fraudArmorCharge = "$0.00";

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - CUSTOMIZE
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            productDetailsPage.Customize.SetQuantity("Duplicates - 2 boxes");
            string productPrice = productDetailsPage.Customize.GetPrice();
            productDetailsPage.Customize.AddLogo(LogoTypes.Stock, 1);
            string logoPrice = productDetailsPage.Customize.GetLogoPrice();
            productDetailsPage.Customize.Continue();

            // PRODUCT DETAILS PAGE - PERSONALIZE
            // TODO: There is an issue with the personalize panel !! Sometimes, the text being partially entered 
            productDetailsPage.Personalize.SetName(customer.name);
            productDetailsPage.Personalize.SetAddressLine(customer.streetAddress);
            productDetailsPage.Personalize.SetCity(customer.city);
            productDetailsPage.Personalize.SetState(customer.state);
            productDetailsPage.Personalize.SetZip(customer.zip);
            productDetailsPage.Personalize.SetOptionalLine(customer.optionalSecondLine);
            productDetailsPage.Personalize.SetFont(1);

            productDetailsPage.Personalize.SetOslLine1("OSL Line 1");
            productDetailsPage.Personalize.SelectSecondSignature();
            string fontPrice = productDetailsPage.Personalize.GetFontPrice();
            string oslPrice = productDetailsPage.Personalize.GetOslPrice();
            productDetailsPage.Personalize.Continue();

            // PRODUCT DETAILS PAGE - ACCOUNT
            productDetailsPage.Account.SetBankName(bankInfo.bankName);
            productDetailsPage.Account.SetStartingCheckNumber(startingCheckNumber);
            productDetailsPage.Account.SetRoutingNumber(account.routingNumber);
            productDetailsPage.Account.SetAccountNumber(account.accountNumber);
            productDetailsPage.Account.SetBottomRowofNumbers(bottomRowOfNumbers);
            productDetailsPage.Account.Next();

            // PRODUCT DETAILS PAGE - FORMAT SELECTION DIALOG
            ProductDetailsPage_FormatSelectionDialog formatSelectionDialog = new ProductDetailsPage_FormatSelectionDialog(Driver);
            formatSelectionDialog.Continue();

            // PRODUCT DETAILS PAGE - CONFIRMATION PANEL
            Assert.IsTrue(productDetailsPage.Confirmation.PreviewImageExists(), "Validate the preview image exists");
            productDetailsPage.Confirmation.SetFraudArmor(false);
            productDetailsPage.Confirmation.ClickVerifyInformationCheckbox();
            productDetailsPage.Confirmation.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(productPrice, shoppingCartPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(logoPrice, shoppingCartPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, shoppingCartPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fontPrice, shoppingCartPage.GetCheckFontPrice(0), "Validate the font price");
            Assert.AreEqual(fraudArmorCharge, shoppingCartPage.GetFraudArmorProtectionPrice(0), "Verify fraud armor price not charged");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChiocePage = new MyAccountChoicePage(Driver);
            myAccountChiocePage.ReturningCheckCustomer(login);

            // BILLING ADDRESS PAGE
            BillingAddressPage billAddressPage = new BillingAddressPage(Driver);
            billAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            shippingDetailsPage.AddAddress();
            shippingDetailsPage.EnterAddress(customer);
            shippingDetailsPage.ShippingAddressContinue();
            shippingDetailsPage.SetShippingMethod(ShippingDetailsPage.ShippingMethods.Priority);

            Assert.AreEqual(productPrice, shippingDetailsPage.GetProductPrice(0), "Validate product price");
            Assert.AreEqual(fontPrice, shippingDetailsPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shippingDetailsPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, shippingDetailsPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, shippingDetailsPage.GetFraudArmorProtectionPrice(0), "Verify fraud armor price not charged");

            Assert.AreEqual(customer.name, shippingDetailsPage.GetLine(1), "Verify customer name");
            Assert.AreEqual(customer.streetAddress, shippingDetailsPage.GetLine(3), "Verify address");
            Assert.AreEqual($"{customer.city}, {customer.state} {customer.zip}", shippingDetailsPage.GetLine(4), "Verify city, state, zip");
            Assert.AreEqual(customer.optionalSecondLine, shippingDetailsPage.GetLine(5), "Validate the optional line");
            Assert.AreEqual(bankInfo.bankName, shippingDetailsPage.GetBankName(), "Verify bank name");
            Assert.AreEqual(account.routingNumber, shippingDetailsPage.GetBankRoutingNumber(), "Verify routing number");
            Assert.AreEqual(account.accountNumber, shippingDetailsPage.GetCheckingAccountNumber(), "Verify checking account number");
            Assert.AreEqual(startingCheckNumber, shippingDetailsPage.GetCheckStartingNumber(), "Verify check starting number");
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            Assert.AreEqual(productPrice, paymentPage.GetProductPrice(0), "Validate product price");
            Assert.AreEqual(fontPrice, paymentPage.GetFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, paymentPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, paymentPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, paymentPage.GetFraudArmorProtectionPrice(0), "Verify fraud armor price not charged");
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(productPrice, orderConfirmationPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, orderConfirmationPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, orderConfirmationPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, orderConfirmationPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, orderConfirmationPage.GetFraudArmorProtectionPrice(0), "Validate fraud armor price not charged");
        }
        [Test, Property("Description", "Validate Business Manual Checks Order - Singles with Fraud Armor and Continental Shipping.")]
        public void TC8043_Business_Manual_Checks_Singles_With_Fraud_Armor()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "102049";
            Account account = new Account("314089681", "0122334455");
            string bottomRowOfNumbers = "31408968101223344551001";
            string startingCheckNumber = "1001";
            Customer customer = Samples.Customers.CONUS;
            BankInfo bankInfo = Samples.Banks.Bank1;

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - CUSTOMIZE
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            productDetailsPage.Customize.SetCheckType(CustomizePanel.CheckTypes.Singles);
            productDetailsPage.Customize.SetQuantity("2 Boxes");
            string productPrice = productDetailsPage.Customize.GetPrice();
            productDetailsPage.Customize.AddLogo(LogoTypes.Stock, 1);
            string logoPrice = productDetailsPage.Customize.GetLogoPrice();
            productDetailsPage.Customize.Continue();

            // PRODUCT DETAILS PAGE - PERSONALIZE
            productDetailsPage.Personalize.SetName(customer.name);
            productDetailsPage.Personalize.SetAddressLine(customer.streetAddress);
            productDetailsPage.Personalize.SetCity(customer.city);
            productDetailsPage.Personalize.SetState(customer.state);
            productDetailsPage.Personalize.SetZip(customer.zip);
            productDetailsPage.Personalize.SetOptionalLine(customer.optionalSecondLine);
            productDetailsPage.Personalize.SetFont(1);
            productDetailsPage.Personalize.SetOslLine1("OSL Line 1");
            productDetailsPage.Personalize.SelectSecondSignature();
            string fontPrice = productDetailsPage.Personalize.GetFontPrice();
            string oslPrice = productDetailsPage.Personalize.GetOslPrice();
            productDetailsPage.Personalize.Continue();

            // PRODUCT DETAILS PAGE - ACCOUNT
            productDetailsPage.Account.SetBankName(bankInfo.bankName);
            productDetailsPage.Account.SetStartingCheckNumber(startingCheckNumber);
            productDetailsPage.Account.SetRoutingNumber(account.routingNumber);
            productDetailsPage.Account.SetAccountNumber(account.accountNumber);
            productDetailsPage.Account.SetBottomRowofNumbers(bottomRowOfNumbers);
            productDetailsPage.Account.Next();

            // PRODUCT DETAILS PAGE - FORMAT SELECTION DIALOG
            ProductDetailsPage_FormatSelectionDialog formatSelectionDialog = new ProductDetailsPage_FormatSelectionDialog(Driver);
            formatSelectionDialog.Continue();

            // CONFIRMATION PANEL
            ConfirmationPanel confirmationPanel = new ConfirmationPanel(Driver);
            confirmationPanel.SetFraudArmor(true); // Passing true will turn on the FraudArmor protection
            // TODO: Cannot validate Fraud Armor Protection because the price is changing, e.g. $0.46 per check on PDP and $20 on cart
            confirmationPanel.ClickVerifyInformationCheckbox();
            confirmationPanel.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(productPrice, shoppingCartPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, shoppingCartPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shoppingCartPage.GetLogoPrice(0), "Validate the Logo price");
            Assert.AreEqual(oslPrice, shoppingCartPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", shoppingCartPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChiocePage = new MyAccountChoicePage(Driver);
            myAccountChiocePage.ContinueAsGuest();

            // BILLING ADDRESS PAGE
            BillingAddressPage billAddressPage = new BillingAddressPage(Driver);
            billAddressPage.EnterAddress(customer);
            billAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            Assert.AreEqual(productPrice, shippingDetailsPage.GetProductPrice(0), "Validate the product selected price");
            Assert.AreEqual(fontPrice, shippingDetailsPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shippingDetailsPage.GetLogoPrice(0), "Validate the Logo price");
            Assert.AreEqual(oslPrice, shippingDetailsPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", shippingDetailsPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
            Assert.AreEqual(customer.name, shippingDetailsPage.GetLine(1), "Verify customer name");
            Assert.AreEqual(customer.streetAddress, shippingDetailsPage.GetLine(3), "Verify address");
            Assert.AreEqual($"{customer.city}, {customer.state} {customer.zip}", shippingDetailsPage.GetLine(4), "Verify city, state, zip");
            Assert.AreEqual(customer.optionalSecondLine, shippingDetailsPage.GetLine(5), "Validate the optional line");
            Assert.AreEqual(bankInfo.bankName, shippingDetailsPage.GetBankName(), "Verify bank name");
            Assert.AreEqual(account.routingNumber, shippingDetailsPage.GetBankRoutingNumber(), "Verify routing number");
            Assert.AreEqual(account.accountNumber, shippingDetailsPage.GetCheckingAccountNumber(), "Verify checking account number");
            Assert.AreEqual(startingCheckNumber, shippingDetailsPage.GetCheckStartingNumber(), "Verify check starting number");
            shippingDetailsPage.SetShippingMethod(ShippingDetailsPage.ShippingMethods.Overnight);
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            Assert.AreEqual(productPrice, paymentPage.GetProductPrice(0), "Validate product price");
            Assert.AreEqual(fontPrice, paymentPage.GetFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, paymentPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, paymentPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", paymentPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(productPrice, orderConfirmationPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, orderConfirmationPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, orderConfirmationPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, orderConfirmationPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", orderConfirmationPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
        }
        [Test, Property("Description", "Validate Business Manual Checks Order - Duplicates without Fraud Armor and Priority shipping.")]
        public void TC8044_Business_Checks_Duplicates_Without_FraudArmor_Priority()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "102051";
            Account account = new Account("314089681", "0122334455");
            string bottomRowOfNumbers = "31408968101223344551001";
            string startingCheckNumber = "1001";
            Customer customer = Samples.Customers.OCONUS2;
            Login login = new Login("susmitha.kolachina@harlandclarke.com", "Testing1");
            BankInfo bankInfo = Samples.Banks.Bank1;
            string fraudArmorCharge = "$0.00";

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - CUSTOMIZE
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            productDetailsPage.Customize.SetQuantity("2 Boxes");
            string productPrice = productDetailsPage.Customize.GetPrice();
            productDetailsPage.Customize.AddLogo(LogoTypes.Stock, 1);
            string logoPrice = productDetailsPage.Customize.GetLogoPrice();
            productDetailsPage.Customize.Continue();

            // PRODUCT DETAILS PAGE - PERSONALIZE
            productDetailsPage = new ProductDetailsPage(Driver);
            productDetailsPage.Personalize.SetName(customer.name);
            productDetailsPage.Personalize.SetAddressLine(customer.streetAddress);
            productDetailsPage.Personalize.SetCity(customer.city);
            productDetailsPage.Personalize.SetState(customer.state);
            productDetailsPage.Personalize.SetZip(customer.zip);
            productDetailsPage.Personalize.SetOptionalLine(customer.optionalSecondLine);
            productDetailsPage.Personalize.SetFont(1);

            productDetailsPage.Personalize.SetOslLine1("OSL Line 1");
            productDetailsPage.Personalize.SelectSecondSignature();
            string fontPrice = productDetailsPage.Personalize.GetFontPrice();
            string oslPrice = productDetailsPage.Personalize.GetOslPrice();
            productDetailsPage.Personalize.Continue();

            // PRODUCT DETAILS PAGE - ACCOUNT
            productDetailsPage = new ProductDetailsPage(Driver);
            productDetailsPage.Account.SetBankName(bankInfo.bankName);
            productDetailsPage.Account.SetStartingCheckNumber(startingCheckNumber);
            productDetailsPage.Account.SetRoutingNumber(account.routingNumber);
            productDetailsPage.Account.SetAccountNumber(account.accountNumber);
            productDetailsPage.Account.SetBottomRowofNumbers(bottomRowOfNumbers);
            productDetailsPage.Account.Next();

            // PRODUCT DETAILS PAGE - FORMAT SELECTION DIALOG
            ProductDetailsPage_FormatSelectionDialog formatSelectionDialog = new ProductDetailsPage_FormatSelectionDialog(Driver);
            formatSelectionDialog.Continue();

            // PRODUCT DETAILS PAGE - CONFIRMATION PANEL
            Assert.IsTrue(productDetailsPage.Confirmation.PreviewImageExists(), "Validate the preview image exists");
            productDetailsPage.Confirmation.SetFraudArmor(false);
            productDetailsPage.Confirmation.ClickVerifyInformationCheckbox();
            productDetailsPage.Confirmation.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(productPrice, shoppingCartPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(logoPrice, shoppingCartPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, shoppingCartPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fontPrice, shoppingCartPage.GetCheckFontPrice(0), "Validate the font price");
            Assert.AreEqual(fraudArmorCharge, shoppingCartPage.GetFraudArmorProtectionPrice(0), "Validate fraud armor price not charged");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChiocePage = new MyAccountChoicePage(Driver);
            myAccountChiocePage.ReturningCheckCustomer(login);

            // BILLING ADDRESS PAGE
            BillingAddressPage billingAddressPage = new BillingAddressPage(Driver);
            billingAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            shippingDetailsPage.AddAddress();
            shippingDetailsPage.EnterAddress(customer);
            shippingDetailsPage.ShippingAddressContinue();
            shippingDetailsPage = new ShippingDetailsPage(Driver);
            shippingDetailsPage.SetShippingMethod(ShippingDetailsPage.ShippingMethods.Priority);
            Assert.AreEqual(productPrice, shippingDetailsPage.GetProductPrice(0), "Validate product price");
            Assert.AreEqual(fontPrice, shippingDetailsPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shippingDetailsPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, shippingDetailsPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, shippingDetailsPage.GetFraudArmorProtectionPrice(0), "Verify fraud armor price not charged");

            Assert.AreEqual(customer.name, shippingDetailsPage.GetLine(1), "Verify customer name");
            Assert.AreEqual(customer.streetAddress, shippingDetailsPage.GetLine(3), "Verify address");
            Assert.AreEqual($"{customer.city}, {customer.state} {customer.zip}", shippingDetailsPage.GetLine(4), "Verify city, state, zip");
            Assert.AreEqual(customer.optionalSecondLine, shippingDetailsPage.GetLine(5), "Validate the optional line");
            Assert.AreEqual(bankInfo.bankName, shippingDetailsPage.GetBankName(), "Verify bank name");
            Assert.AreEqual(account.routingNumber, shippingDetailsPage.GetBankRoutingNumber(), "Verify routing number");
            Assert.AreEqual(account.accountNumber, shippingDetailsPage.GetCheckingAccountNumber(), "Verify checking account number");
            Assert.AreEqual(startingCheckNumber, shippingDetailsPage.GetCheckStartingNumber(), "Verify check starting number");
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            Assert.AreEqual(productPrice, paymentPage.GetProductPrice(0), "Validate product price");
            Assert.AreEqual(fontPrice, paymentPage.GetFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, paymentPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, paymentPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, paymentPage.GetFraudArmorProtectionPrice(0), "Verify fraud armor price not charged");
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(productPrice, orderConfirmationPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, orderConfirmationPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, orderConfirmationPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, orderConfirmationPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, orderConfirmationPage.GetFraudArmorProtectionPrice(0), "Validate fraud armor price not charged");
        }
        [Test, Property("Description", "Validate Business Computer Checks Order without Fraud Armor and priority shipping")]
        public void TC8045_Business_Computer_Checks_Duplicates_Without_FraudArmor_Priority()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "101872";
            Account account = new Account("314089681", "0122334455");
            string bottomRowOfNumbers = "31408968101223344551001";
            string startingCheckNumber = "1001";
            Customer customer = Samples.Customers.OCONUS2;
            Login login = new Login("susmitha.kolachina@harlandclarke.com", "Testing1");
            BankInfo bankInfo = Samples.Banks.Bank1;
            string fraudArmorCharge = "$0.00";

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - CUSTOMIZE
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            productDetailsPage.Customize.SetQuantity("500");
            string productPrice = productDetailsPage.Customize.GetPrice();
            productDetailsPage.Customize.SetStyleColor(ProductStyleColors.Blue);
            productDetailsPage.Customize.AddLogo(LogoTypes.Stock, 1);
            string logoPrice = productDetailsPage.Customize.GetLogoPrice();
            productDetailsPage.Customize.Continue();

            // PRODUCT DETAILS PAGE - PERSONALIZE
            // TODO: There is an issue with the personalize panel !! Sometimes, the text being partially entered 
            productDetailsPage = new ProductDetailsPage(Driver);
            productDetailsPage.Personalize.SetName(customer.name);
            productDetailsPage.Personalize.SetAddressLine(customer.streetAddress);
            productDetailsPage.Personalize.SetCity(customer.city);
            productDetailsPage.Personalize.SetState(customer.state);
            productDetailsPage.Personalize.SetZip(customer.zip);
            productDetailsPage.Personalize.SetOptionalLine(customer.optionalSecondLine);
            productDetailsPage.Personalize.SetFont(1);

            productDetailsPage.Personalize.SetOslLine1("OSL Line 1");
            productDetailsPage.Personalize.SelectSecondSignature();
            string fontPrice = productDetailsPage.Personalize.GetFontPrice();
            string oslPrice = productDetailsPage.Personalize.GetOslPrice();
            productDetailsPage.Personalize.Continue();

            // PRODUCT DETAILS PAGE - ACCOUNT
            productDetailsPage = new ProductDetailsPage(Driver);
            productDetailsPage.Account.SetBankName(bankInfo.bankName);
            productDetailsPage.Account.SetStartingCheckNumber(startingCheckNumber);
            productDetailsPage.Account.SetRoutingNumber(account.routingNumber);
            productDetailsPage.Account.SetAccountNumber(account.accountNumber);
            productDetailsPage.Account.SetBottomRowofNumbers(bottomRowOfNumbers);
            productDetailsPage.Account.Next();

            // PRODUCT DETAILS PAGE - FORMAT SELECTION DIALOG
            ProductDetailsPage_FormatSelectionDialog formatSelectionDialog = new ProductDetailsPage_FormatSelectionDialog(Driver);
            formatSelectionDialog.Continue();

            // PRODUCT DETAILS PAGE - CONFIRMATION PANEL
            Assert.IsTrue(productDetailsPage.Confirmation.PreviewImageExists(), "Validate the preview image exists");
            productDetailsPage.Confirmation.SetFraudArmor(false);
            productDetailsPage.Confirmation.ClickVerifyInformationCheckbox();
            productDetailsPage.Confirmation.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(productPrice, shoppingCartPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(logoPrice, shoppingCartPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, shoppingCartPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fontPrice, shoppingCartPage.GetCheckFontPrice(0), "Validate the font price");
            Assert.AreEqual(fraudArmorCharge, shoppingCartPage.GetFraudArmorProtectionPrice(0), "Verify fraud armor price not charged");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChiocePage = new MyAccountChoicePage(Driver);
            myAccountChiocePage.ReturningCheckCustomer(login);

            // BILLING ADDRESS PAGE
            BillingAddressPage billAddressPage = new BillingAddressPage(Driver);
            billAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            shippingDetailsPage.AddAddress();
            shippingDetailsPage.EnterAddress(customer);
            shippingDetailsPage.ShippingAddressContinue();
            shippingDetailsPage = new ShippingDetailsPage(Driver);
            shippingDetailsPage.SetShippingMethod(ShippingDetailsPage.ShippingMethods.Priority);

            Assert.AreEqual(productPrice, shippingDetailsPage.GetProductPrice(0), "Validate product price");
            Assert.AreEqual(fontPrice, shippingDetailsPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shippingDetailsPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, shippingDetailsPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, shippingDetailsPage.GetFraudArmorProtectionPrice(0), "Verify fraud armor price not charged");

            Assert.AreEqual(customer.name, shippingDetailsPage.GetLine(1), "Verify customer name");
            Assert.AreEqual(customer.streetAddress, shippingDetailsPage.GetLine(3), "Verify address");
            Assert.AreEqual($"{customer.city}, {customer.state} {customer.zip}", shippingDetailsPage.GetLine(4), "Verify city, state, zip");
            Assert.AreEqual(customer.optionalSecondLine, shippingDetailsPage.GetLine(5), "Validate the optional line");
            Assert.AreEqual(bankInfo.bankName, shippingDetailsPage.GetBankName(), "Verify bank name");
            Assert.AreEqual(account.routingNumber, shippingDetailsPage.GetBankRoutingNumber(), "Verify routing number");
            Assert.AreEqual(account.accountNumber, shippingDetailsPage.GetCheckingAccountNumber(), "Verify checking account number");
            Assert.AreEqual(startingCheckNumber, shippingDetailsPage.GetCheckStartingNumber(), "Verify check starting number");
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            Assert.AreEqual(productPrice, paymentPage.GetProductPrice(0), "Validate product price");
            Assert.AreEqual(fontPrice, paymentPage.GetFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, paymentPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, paymentPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, paymentPage.GetFraudArmorProtectionPrice(0), "Verify fraud armor price not charged");
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(productPrice, orderConfirmationPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, orderConfirmationPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, orderConfirmationPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, orderConfirmationPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreEqual(fraudArmorCharge, orderConfirmationPage.GetFraudArmorProtectionPrice(0), "Validate fraud armor price not charged");
        }
        [Test, Property("Description", "Validate Business Computer Checks Order with Fraud Armor and PO Box shipping validation")]
        public void TC8046_Business_Computer_Checks_Singles_With_Fraud_Armor()
        {
            BrowseablePages.init(site, subDomain);

            // TEST CASE DATA
            string productId = "98540";
            Account account = new Account("314089681", "0122334455");
            string bottomRowOfNumbers = "31408968101223344551001";
            string startingCheckNumber = "1001";
            Customer pOBoxCustomer = new Customer("testaccount2@citm.com", false, "1234567890", "John CONUS", "HC Test Order", true, "PO BOX 7530", "", "San Antonio", "TX", "78256", "Optional phone line");
            Customer customer = Samples.Customers.CONUS;
            BankInfo bankInfo = Samples.Banks.Bank1;
            string errorMessage = "For security reasons, business checks cannot be shipped to a PO Box address. Please specify another address.";

            Driver.Manage().Window.Maximize();
            Driver.Url = BrowseablePages.domain + "/p/" + productId;

            // CHECK FOR SECURITY PAGE IN IE
            if (browser == Browsers.IE)
            {
                Driver.FindElement(By.Id("overridelink")).Click();
            }

            // PRODUCT DETAILS PAGE - CUSTOMIZE
            ProductDetailsPage productDetailsPage = new ProductDetailsPage(Driver);
            Assert.IsTrue(productDetailsPage.Customize.PreviewImageExists(), "Validate preview image exists");
            productDetailsPage.Customize.SetQuantity("500");
            string productPrice = productDetailsPage.Customize.GetPrice();
            productDetailsPage.Customize.SetStyleColor(ProductStyleColors.Antique);
            productDetailsPage.Customize.AddLogo(LogoTypes.Stock, 1);
            string logoPrice = productDetailsPage.Customize.GetLogoPrice();
            productDetailsPage.Customize.Continue();

            // PRODUCT DETAILS PAGE - PERSONALIZE
            productDetailsPage = new ProductDetailsPage(Driver);
            productDetailsPage.Personalize.SetName(customer.name);
            productDetailsPage.Personalize.SetAddressLine(customer.streetAddress);
            productDetailsPage.Personalize.SetCity(customer.city);
            productDetailsPage.Personalize.SetState(customer.state);
            productDetailsPage.Personalize.SetZip(customer.zip);
            productDetailsPage.Personalize.SetOptionalLine(customer.optionalSecondLine);
            productDetailsPage.Personalize.SetFont(1);
            productDetailsPage.Personalize.SetOslLine1("OSL Line 1");
            productDetailsPage.Personalize.SelectSecondSignature();
            string fontPrice = productDetailsPage.Personalize.GetFontPrice();
            string oslPrice = productDetailsPage.Personalize.GetOslPrice();
            productDetailsPage.Personalize.Continue();

            // PRODUCT DETAILS PAGE - ACCOUNT
            productDetailsPage = new ProductDetailsPage(Driver);
            productDetailsPage.Account.SetBankName(bankInfo.bankName);
            productDetailsPage.Account.SetStartingCheckNumber(startingCheckNumber);
            productDetailsPage.Account.SetRoutingNumber(account.routingNumber);
            productDetailsPage.Account.SetAccountNumber(account.accountNumber);
            productDetailsPage.Account.SetBottomRowofNumbers(bottomRowOfNumbers);
            productDetailsPage.Account.Next();

            // PRODUCT DETAILS PAGE - FORMAT SELECTION DIALOG
            ProductDetailsPage_FormatSelectionDialog formatSelectionDialog = new ProductDetailsPage_FormatSelectionDialog(Driver);
            formatSelectionDialog.Continue();

            // CONFIRMATION PANEL
            ConfirmationPanel confirmationPanel = new ConfirmationPanel(Driver);
            confirmationPanel.SetFraudArmor(true); // Passing true will turn on the FraudArmor protection
            // TODO: Cannot validate Fraud Armor Protection because the price is changing, e.g. $0.46 per check on PDP and $20 on cart
            confirmationPanel.ClickVerifyInformationCheckbox();
            confirmationPanel.AddToCart();

            // SHOPPING CART PAGE
            ShoppingCartPage shoppingCartPage = new ShoppingCartPage(Driver);
            Assert.AreEqual(productPrice, shoppingCartPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, shoppingCartPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shoppingCartPage.GetLogoPrice(0), "Validate the Logo price");
            Assert.AreEqual(oslPrice, shoppingCartPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", shoppingCartPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
            shoppingCartPage.ProceedToCheckout();

            // MY ACCOUNT CHOICE PAGE
            MyAccountChoicePage myAccountChiocePage = new MyAccountChoicePage(Driver);
            myAccountChiocePage.ContinueAsGuest();

            // BILLING ADDRESS PAGE
            BillingAddressPage billAddressPage = new BillingAddressPage(Driver);
            billAddressPage.EnterAddress(pOBoxCustomer);
            billAddressPage.Continue();

            // SHIPPING DETAILS PAGE
            ShippingDetailsPage shippingDetailsPage = new ShippingDetailsPage(Driver);
            Assert.AreEqual(productPrice, shippingDetailsPage.GetProductPrice(0), "Validate the product selected price");
            Assert.AreEqual(fontPrice, shippingDetailsPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, shippingDetailsPage.GetLogoPrice(0), "Validate the Logo price");
            Assert.AreEqual(oslPrice, shippingDetailsPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", shippingDetailsPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");

            Assert.AreEqual(customer.name, shippingDetailsPage.GetLine(1), "Verify customer name");
            Assert.AreEqual(customer.streetAddress, shippingDetailsPage.GetLine(3), "Verify address");
            Assert.AreEqual($"{customer.city}, {customer.state} {customer.zip}", shippingDetailsPage.GetLine(4), "Verify city, state, zip");
            Assert.AreEqual(customer.optionalSecondLine, shippingDetailsPage.GetLine(5), "Validate the optional line");
            Assert.AreEqual(bankInfo.bankName, shippingDetailsPage.GetBankName(), "Verify bank name");
            Assert.AreEqual(account.routingNumber, shippingDetailsPage.GetBankRoutingNumber(), "Verify routing number");
            Assert.AreEqual(account.accountNumber, shippingDetailsPage.GetCheckingAccountNumber(), "Verify checking account number");
            Assert.AreEqual(startingCheckNumber, shippingDetailsPage.GetCheckStartingNumber(), "Verify check starting number");
            Assert.AreEqual(errorMessage, shippingDetailsPage.GetErrorMessage(), "Validate that error message is shown");
            shippingDetailsPage.EditAddress();
            shippingDetailsPage.EnterAddress(customer);
            shippingDetailsPage.UpdateShippingAddress();
            shippingDetailsPage = new ShippingDetailsPage(Driver);
            Assert.AreEqual(string.Empty, shippingDetailsPage.GetErrorMessage(), "Validate that error message not displayed");
            shippingDetailsPage.SetShippingMethod(ShippingDetailsPage.ShippingMethods.Overnight);
            shippingDetailsPage.Continue();

            // PAYMENT PAGE
            PaymentPage paymentPage = new PaymentPage(Driver);
            Assert.AreEqual(productPrice, paymentPage.GetProductPrice(0), "Validate product price");
            Assert.AreEqual(fontPrice, paymentPage.GetFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, paymentPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, paymentPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", paymentPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
            paymentPage.SubmitOrder();

            // ORDER CONFIRMATION PAGE
            OrderConfirmationPage orderConfirmationPage = new OrderConfirmationPage(Driver);
            Assert.AreEqual(productPrice, orderConfirmationPage.GetProductPrice(0), "Validate the product price");
            Assert.AreEqual(fontPrice, orderConfirmationPage.GetCheckFontPrice(0), "Validate font price");
            Assert.AreEqual(logoPrice, orderConfirmationPage.GetLogoPrice(0), "Validate the logo price");
            Assert.AreEqual(oslPrice, orderConfirmationPage.GetOslPrice(0), "Validate the osl price");
            Assert.AreNotEqual("$0.00", orderConfirmationPage.GetFraudArmorProtectionPrice(0), "Validate that fraud armor is off");
        }
        [TearDown]
        public void TearDown()
        {
            if (!localRun)
            {
                Functions.LogResult(Driver);
            }
            Driver.Quit();
        }
    }
}