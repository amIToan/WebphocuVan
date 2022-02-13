
; (function ($) {

    'use strict'

    var scrollCalc = function () {
        if ($(window).scrollTop() > 1) {
            $('#header').addClass('float-header');
            $('.header').addClass('float-header');

            $("#selectLang").css("margin-top", "0px");
            if ($(window).innerWidth() > 992) {
            } else {
                $(".btn-menu").css("margin-top", "0px");
            }
        } else {
            $('#header').removeClass('float-header');
            $('.header').removeClass('float-header');

            if ($(window).innerWidth() > 992) {
                $("#selectLang").css("margin-top", "12px");
            } else {
                $(".btn-menu").css("margin-top", "36px");
            }
        }
    };

    var init_header = function () {
        var largeScreen = matchMedia('only screen and (min-width: 992px)').matches;
        if (largeScreen) {
            if ($().sticky) {
                $('header.header-sticky').sticky();
            }
        }

        $('.one-page .mainnav ul > li > a').on('click', function () {
            var anchor = $(this).attr('href').split('#')[1];
            var headerHeight = 0;
            var headerSmall = 80;

            if ($(window).scrollTop() != 0) {
                headerSmall = 0;
            }

            if (anchor) {
                if ($('#' + anchor).length > 0) {
                    if ($('.header-sticky').length > 0 && largeScreen) {
                        headerHeight = $('.header').outerHeight();
                    }
                    var target = $('.' + anchor).offset().top - headerHeight + headerSmall + 1;
                    $('html,body').animate({ scrollTop: target }, 1000, 'easeInOutExpo');
                }
            }
            //return false;
        }); // click on one-page menu

        $('.one-page .mainnav > ul > li > a').on('click', function () {
            $(this).addClass('active').parent().siblings().children().removeClass('active');
        });
    };

    var fullScreen = function () {
        (function () {
            function setupSlider() {
                if ($('body').hasClass('full-screen')) {
                    $('.control-slider').css({ 'margin-top': 0, 'padding-bottom': 0 }); //reset slider

                    var wHeight = $(window).height(),
                        contentHeight = $('.control-slider').outerHeight(),
                        contactHeight = $('.contact-slider').outerHeight(),
                        contactBottom = parseInt($('.contact-slider').css('bottom')),
                        hHeader = $('.sticky-wrapper').outerHeight() + $('.sticky-wrapper').outerHeight() / 2,
                        margin = (wHeight - contentHeight - contactHeight - hHeader) / 2;

                    $('.control-slider').css({ 'margin-top': margin, 'padding-bottom': margin + $('.sticky-wrapper').outerHeight() / 2 })
                }
            }

            $(window).on("resize", setupSlider);
            $(document).on("ready", setupSlider);
        })(); // set fullscreen and vertical align for content

        (function () {
            $('.btn-top').on('click', function () {
                var anchor = $(this).attr('href').split('#')[1];
                if (anchor) {
                    if ($('#' + anchor).length > 0) {
                        var headerHeight = 0;
                        if ($('.header-sticky').length > 0) {
                            headerHeight = $('.header').outerHeight();
                        }
                        var target = $('#' + anchor).offset().top - headerHeight;

                        $('html,body').animate({ scrollTop: target }, 1000, 'easeInOutExpo');
                    }
                }
                return false;
            });
        })();
    };

    var controlSlider = function () {
        if ($().flexslider) {
            $('.control-slider').each(function () {
                $(this).find('.flexslider').flexslider({
                    animation: 'fade',
                    animationLoop: true,
                    slideshow: true,
                    slideshowSpeed: 6000,
                    animationSpeed: 800,
                    pauseOnHover: true,
                    pauseOnAction: true,
                    controlNav: true,
                    directionNav: false,
                    prevText: '<i class="fa fa-angle-left"></i>',
                    nextText: '<i class="fa fa-angle-right"></i>',
                    controlsContainer: '.flex-container',
                    start: function () {
                        $(".control-slider").css("margin-top", "0px");
                        // $(".content-slider").css("margin-top", "200px");
                        $(".flex-control-nav").css("top", ($(window).innerHeight() - 80 - $(".contact-slider").height()) + "px");
                        $(".flex-control-nav").css("z-index", "2");
                        if ($(window).innerHeight() < 768) {
                            $(".content-slider").css("margin-top", "200px");

                        } else {
                            $(".content-slider").css("margin-top", "200px");
                            //   $(".content-slider").css("margin-top", "200px");

                        }
                        //alert( )
                    },
                    before: function (e) {
                        //   alert("test");

                        //  alert(this);
                    }
                }); // end flexslider
            }); // end each
        } // end if
    };

    var retinaLogos = function () {
        var retina = window.devicePixelRatio > 1 ? true : false;

        if (retina) {
            $('.logo img').attr({ src: './images/logo.png', width: '169', height: '35' });
        }

        if (retina) {
            $('.h-logo img').attr({ src: '/images/logo.png', width: '169', height: '35' });
        }
    };

    var galleryIsotope = function () {
        if ($().isotope) {
            var $container = $('.gallery-wrap');

            $container.imagesLoaded(function () {
                $container.isotope({
                    itemSelector: '.gallery-item',
                    transitionDuration: '1s'
                }); // end isotope
            });

            $('.gallery-filter li').on('click', function () {
                $('.gallery-filter li').removeClass('active');
                $(this).addClass('active');
                var selector = $(this).find("a").attr('data-filter');
                $container.isotope({ filter: selector });
                return false;
            }); // end filter
        };
    };

    var foodIsotope = function () {
        if ($().isotope) {
            var $container = $('.food-wrap');

            $container.imagesLoaded(function () {
                $container.isotope({
                    itemSelector: '.food-item',
                    transitionDuration: '0.6s',
                    hiddenStyle: { opacity: 0 },
                    visibleStyle: { opacity: 1 }
                }); // end isotope
            });

            $('.food-filter li').on('click', function () {
                var selector = $(this).find("a").attr('data-filter')+"";
                
                $('.FoodMenu').fadeOut();


                
                $(selector).fadeIn();
                $('.food-filter li').removeClass('active');
                $(this).addClass('active');
               // $container.isotope({ filter: selector });

                return false;
            }); // end filter

            $('.food-filter li:nth-child(5)').trigger('click');
        };
    };

    var blogMasonry = function () {
        if ($().isotope) {
            var $container = $('.posts-container');

            $container.imagesLoaded(function () {
                $container.isotope({
                    itemSelector: '.post-item',
                    transitionDuration: '0.5s',
                    layoutMode: 'masonry',
                    masonry: { columnWidth: $container.width() / 12 }
                }); // isotope
            });

            $(window).resize(function () {
                $container.isotope({
                    masonry: { columnWidth: $container.width() / 12 }
                });
            }); // relayout

            $('.post-filter li').on('click', function () {
                $('.post-filter li').removeClass('active');
                $(this).addClass('active');
                var selector = $(this).find("a").attr('data-filter');
                $container.isotope({ filter: selector });
                return false;
            }); // filter
        };
    };

    var testimonial = function () {
        $('.testimonial-text').bxSlider({
            mode: 'fade',
            touchEnabled: true,
            oneToOneTouch: true,
            pagerCustom: '#bx-pager',
            nextSelector: '#bx-next',
            prevSelector: '#bx-prev',
            nextText: '<i class="fa fa-angle-right"></i>',
            prevText: '<i class="fa fa-angle-left"></i>'
        });
    };

    var simpleSlider = function () {
        if ($().flexslider) {
            $('.roll-slider').each(function () {
                var $this = $(this);
                var easing = ($this.data('effect') == 'fade') ? 'linear' : 'easeInOutExpo';
                $this.find('.flexslider').flexslider({
                    animation: $this.data('effect'),
                    direction: $this.data('direction'), // vertical
                    pauseOnHover: true,
                    useCSS: false,
                    easing: easing,
                    animationSpeed: 500,
                    slideshowSpeed: 5000,
                    controlNav: true,
                    directionNav: true,
                    smoothHeight: true,
                    slideshow: $this.data('auto'),
                    prevText: '<i class="fa fa-angle-left"></i>',
                    nextText: '<i class="fa fa-angle-right"></i>'
                }); // end flexslider
            }); // end roll-slider
        }
    };

    var blogSlider = function () {
        if ($().owlCarousel) {
            $(".roll-blog-slider").owlCarousel({
                navigation: false,
                pagination: false,
                responsive: true,
                items: 1,
                itemsDesktop: [3000, 4],
                itemsDesktopSmall: [1400, 4],
                itemsTablet: [970, 2],
                itemsTabletSmall: [600, 2],
                itemsMobile: [360, 1],
                touchDrag: true,
                mouseDrag: true,
                autoHeight: true,
                autoPlay: false
            }); // end owlCarousel
        }
    };

    var toggles = function () {
        var args = { easing: 'easeOutExpo', duration: 600 };
        $('.toggle .toggle-title.active').siblings('.toggle-content').show();

        $('.toggle.toggle-enable .toggle-title').on('click', function () {
            $(this).closest('.toggle').find('.toggle-content').slideToggle(args);
            $(this).toggleClass('active');
        }); // toggle 

        $('.accordion .toggle-title').on('click', function () {
            if (!$(this).is('.active')) {
                $(this).closest('.accordion').find('.toggle-title.active').toggleClass('active').next().slideToggle(args);
                $(this).toggleClass('active');
                $(this).next().slideToggle(args);
            } else {
                $(this).toggleClass('active');
                $(this).next().slideToggle(args);
            }
        }); // accordion
    };

    var progressBar = function () {
        $('.progress-bar').on('on-appear', function () {
            $(this).each(function () {
                var percent = $(this).data('percent');

                $(this).find('.progress-animate').animate({
                    "width": percent + '%'
                }, 3000);

                $(this).parent('.progress-single').find('.perc').addClass('show').animate({
                    "width": percent + '%'
                }, 3000);
            });
        });
    }

    var ResponsiveMenu = {

        menuType: 'desktop',

        initial: function (winWidth) {
            ResponsiveMenu.menuWidthDetect(winWidth);
            ResponsiveMenu.menuBtnClick();
            ResponsiveMenu.parentMenuClick();
        },

        menuWidthDetect: function (winWidth) {
            var currMenuType = 'desktop';

            if (matchMedia('only screen and (max-width: 978px)').matches) {
                currMenuType = 'mobile';
            } // change menu type

            if (currMenuType !== ResponsiveMenu.menuType) {
                ResponsiveMenu.menuType = currMenuType;

                if (currMenuType === 'mobile') {
                    var $mobileMenu = $('.mainnav').attr('class', 'mainnav-mobi').hide();
                    $('.header').find('.header-wrap').after($mobileMenu);
                    var hasChildMenu = $('.mainnav-mobi').find('li:has(ul)');
                    hasChildMenu.children('ul').hide();
                    hasChildMenu.children('a').after('<span class="btn-submenu"></span>');
                    $('.btn-menu').removeClass('active');
                } else {
                    var $desktopMenu = $('.mainnav-mobi').attr('class', 'mainnav').removeAttr('style');
                    $desktopMenu.find('.sub-menu').removeAttr('style');
                    $('.header').find('.span10').after($desktopMenu);
                    $('.btn-submenu').remove();
                }
            } // clone and insert menu
        },

        menuBtnClick: function () {
            $('.btn-menu').on('click', function () {
                $('.mainnav-mobi').slideToggle(300);
                $(this).toggleClass('active');
            });
        }, // click on moblie button

        parentMenuClick: function () {
            $(document).on('click', '.mainnav-mobi li .btn-submenu', function (e) {
                if ($(this).has('ul')) {
                    e.stopImmediatePropagation()
                    $(this).next('ul').slideToggle(300);
                    $(this).toggleClass('active');
                }
            });
        } // click on sub-menu button
    };

    var ajaxSubscribe = {
        obj: {
            //subscribeEmail: $('#subscribe-email'),
            //subscribeButton: $('#subscribe-button'),
            //subscribeMsg: $('#subscribe-msg'),
            //subscribeContent: $("#subscribe-content"),
            //dataMailchimp: $('#subscribe-form').attr('data-mailchimp'),
            //success_message: '<div class="notification_ok">Thank you for joining our mailing list.<br>Please check your email for a confirmation link.</div>',
            //failure_message: '<div class="notification_error">There was a problem processing your submission.</div>',
            //noticeError: '<div class="notification_error">{msg}</div>',
            //noticeInfo: '<div class="notification_error">{msg}</div>',
            //basicAction: 'mail/subscribe.php',
            //mailChimpAction: 'mail/subscribe-mailchimp.php'
        },

        eventLoad: function () {
            var objUse = ajaxSubscribe.obj;
            $(objUse.subscribeButton).on('click', function () {
                if (window.ajaxCalling) return;
                var isMailchimp = objUse.dataMailchimp === 'true';

                if (isMailchimp) {
                    ajaxSubscribe.ajaxCall(objUse.mailChimpAction);
                } else {
                    ajaxSubscribe.ajaxCall(objUse.basicAction);
                }
            });
        },

        ajaxCall: function (action) {
            window.ajaxCalling = true;
            var objUse = ajaxSubscribe.obj;
            var messageDiv = objUse.subscribeMsg.html('').hide();
            $.ajax({
                url: action,
                type: 'POST',
                dataType: 'json',
                data: {
                    subscribeEmail: objUse.subscribeEmail.val()
                },
                success: function (responseData, textStatus, jqXHR) {
                    if (responseData.status) {
                        objUse.subscribeContent.fadeOut(500, function () {
                            messageDiv.html(objUse.success_message).fadeIn(500);
                        });
                    } else {
                        switch (responseData.msg) {
                            case "email-required":
                                messageDiv.html(objUse.noticeError.replace('{msg}', 'Email is required.'));
                                break;
                            case "email-err":
                                messageDiv.html(objUse.noticeError.replace('{msg}', 'Email invalid.'));
                                break;
                            case "duplicate":
                                messageDiv.html(objUse.noticeError.replace('{msg}', 'Email is duplicate.'));
                                break;
                            case "filewrite":
                                messageDiv.html(objUse.noticeInfo.replace('{msg}', 'Mail list file is open.'));
                                break;
                            case "undefined":
                                messageDiv.html(objUse.noticeInfo.replace('{msg}', 'undefined error.'));
                                break;
                            case "api-error":
                                objUse.subscribeContent.fadeOut(500, function () {
                                    messageDiv.html(objUse.failure_message);
                                });
                        }
                        messageDiv.fadeIn(500);
                    }
                },
                error: function (jqXHR, textStatus, errorThrown) {
                    alert('Connection error');
                },
                complete: function (data) {
                    window.ajaxCalling = false;
                }
            });
        }
    };

    var orAnimation = function () {
        $('.orches-animation').each(function () {
            var orElement = $(this),
               orAnimationClass = orElement.data('animation'),
               orAnimationDelay = orElement.data('animation-delay'),
               orAnimationOffset = orElement.data('animation-offset');

            orElement.css({
                '-webkit-animation-delay': orAnimationDelay,
                '-moz-animation-delay': orAnimationDelay,
                'animation-delay': orAnimationDelay
            });

            orElement.waypoint(function () {
                orElement.addClass('animated').addClass(orAnimationClass);
            }, {
                triggerOnce: true,
                offset: orAnimationOffset
            });
        });
    };

    var effectGallery = function () {
        var effect = $('.gallery-wrap').data('portfolio-effect');
        $('.gallery-item').children('.item-wrap').addClass('orches-animation');

        $('.gallery-wrap').waypoint(function (direction) {
            $('.gallery-item').children('.item-wrap').each(function (idx, ele) {
                setTimeout(function () {
                    $(ele).addClass('animated ' + effect);
                }, idx * 150);
            });
        }, {
            offset: '75%'
        });
    };

    var goTop = function () {
        $(window).scroll(function () {
            if ($(this).scrollTop() > 800) {
                $('.go-top').addClass('show');
            } else {
                $('.go-top').removeClass('show');
            }
        });

        $('.go-top').on('click', function () {
            $("html, body").animate({ scrollTop: 0 }, 1000, 'easeInOutExpo');
            return false;
        });
    };

    var lastestTweets = function () {
        if ($().tweet) {
            $('.latest-tweets').each(function () {
                var $this = $(this);
                $this.tweet({
                    username: $this.data('username'),
                    join_text: "auto",
                    avatar_size: null,
                    count: $this.data('number'),
                    template: "{text}",
                    loading_text: "loading tweets...",
                    modpath: $this.data('modpath')
                }); // tweet
            }); // lastest-tweets each
        }
    };


    var flickrFeed = function () {
        if ($().jflickrfeed) {
            $('.flickr-photos').each(function () {
                $(this).jflickrfeed({
                    limit: 12,
                    qstrings: { id: '130700496@N03' }, // Your Flickr Id 
                    itemTemplate: '<li class="thumb"><span class="overlay-pop"></span><a href="{{link}}" title="{{title}}" target="_blank"><img src="{{image_s}}" alt="{{title}}" /></a></li>'
                });
            });
        }
    };

    var orCountdown = function () {
        var zupa_style = function (data) {
            $(this.el).html(
               "<div class='column days'>" +
                   "<div class='text'>Days</div>" +
                   "<div class='numb'>" + this.leadingZeros(data.days, 2) + "</div>" +
               "</div>" +
               "<div class='column hours'>" +
                   "<div class='text'>Hours</div>" +
                   "<div class='numb'>" + this.leadingZeros(data.hours, 2) + "</div>" +
               "</div>" +
               "<div class='column mins'>" +
                   "<div class='text'>Minutes</div>" +
                   "<div class='numb'>" + this.leadingZeros(data.min, 2) + "</div>" +
               "</div>" +
               "<div class='column secs'>" +
                   "<div class='text'>Seconds</div>" +
                   "<div class='numb'>" + this.leadingZeros(data.sec, 2) + "</div>" +
               "</div>");
        }

        $('.countdown').each(function () {
            $(this).countdown({
                date: $(this).attr('data-date'),
                render: zupa_style
            });
        });
    };

    var parallax = function () {
        $('.parallax-bg1').parallax("50%", 0.5);
        $('.parallax-bg2').parallax("50%", 0.5);
        $('.parallax-bg4').parallax("50%", 0.3);
        $('.parallax-bg5').parallax("50%", 0.5);
        $('.parallax-bg7').parallax("50%", 0.7);
        $('.parallax-bg9').parallax("50%", 0.4);
        $('.parallax-bg12').parallax("50%", 0.4);
    };

    var teamCarousel = function () {
        if ($().owlCarousel) {
            $(".roll-team").owlCarousel({
                navigation: false,
                pagination: true,
                responsive: true,
                items: 1,
                itemsDesktop: [3000, 4],
                itemsDesktopSmall: [1400, 4],
                itemsTablet: [970, 2],
                itemsTabletSmall: [600, 2],
                itemsMobile: [360, 1],
                touchDrag: true,
                mouseDrag: true,
                autoHeight: true,
                autoPlay: false
            }); // end owlCarousel
        } // end if
    };

    var testiCarousel = function () {
        if ($().owlCarousel) {
            $(".roll-testimonials").owlCarousel({
                navigation: false,
                pagination: true,
                responsive: true,
                items: 1,
                itemsDesktop: [3000, 1],
                itemsDesktopSmall: [1400, 1],
                itemsTablet: [970, 1],
                itemsTabletSmall: [600, 1],
                itemsMobile: [360, 1],
                touchDrag: true,
                mouseDrag: true,
                autoHeight: true,
                autoPlay: false
            }); // end owlCarousel
        } // end if
    };

    var googleMap = function () {
        if ($().gmap3) {
            $("#map").gmap3({
                map: {
                    options: {
                        zoom: 14,
                        mapTypeId: 'zupa_style',
                        mapTypeControlOptions: {
                            mapTypeIds: ['zupa_style', google.maps.MapTypeId.SATELLITE, google.maps.MapTypeId.HYBRID]
                        },
                        scrollwheel: false
                    }
                },
                getlatlng: {
                    address: "3 London Rd London SE1 6JZ United Kingdom",
                    callback: function (results) {
                        if (!results) return;
                        $(this).gmap3('get').setCenter(new google.maps.LatLng(results[0].geometry.location.lat(), results[0].geometry.location.lng()));
                        $(this).gmap3({
                            marker: {
                                latLng: results[0].geometry.location,
                                options: {
                                    icon: 'http://s18.postimg.org/rylvem679/zupa_map.png'
                                }
                            }
                        });
                    }
                },
                styledmaptype: {
                    id: "zupa_style",
                    options: {
                        name: "Zupa Map"
                    },
                },
            });
        }
    };

    var customVideo = function () {
        if ($().fitVids)
            $('.blog-post .post .feature-post').fitVids();
    }

    var topVideoBg = function () {
        if ($().YTPlayer) {
            $(".top-slider.video-bg .player").YTPlayer({
                onReady: function (player) {
                    $('.top-slider .load-video').remove();
                }
            });
            if ($().fitVids)
                $('.top-slider').fitVids();
        }
    }

    // Dom Ready
    $(function () {
        init_header();
        fullScreen();
        simpleSlider();
        controlSlider();
        foodIsotope();
        teamCarousel();
        customVideo();
        topVideoBg();

        $(window).scroll(function () { scrollCalc(); })
        $(document).ready(function () { scrollCalc(); })

        $("#wi-top-area.type-fullscreen-slider .super").each(function () {
            var $this = $(this);
            var number = $this.data('number');
            number = parseInt(number);

            var i = 1;
            var slides = [];
            while (i <= number) {
                slides.push({ image: $this.data('image-' + i) });
                i++;
            }
            $this.supersized({
                autoplay: true,
                slide_interval: 4000,
                transition: $this.data('effect'),
                transition_speed: 700,
                slides: slides,
                fit_portrait: false,
            }); // supersized
        });

        $('.datepick').on('click', function () {
            $(this).siblings('input').focus();
            return false;
        });

        retinaLogos();
        galleryIsotope();
        blogMasonry();
        testimonial();
        toggles();
        progressBar();
        orAnimation();
        effectGallery();
        goTop()
        orCountdown();
        ajaxSubscribe.eventLoad();
        parallax();
        // Initialize responsive menu
        ResponsiveMenu.initial($(window).width());
        $(window).resize(function () {
            ResponsiveMenu.menuWidthDetect($(this).width());
        });

        // Detect elements into viewport
        $('[data-waypoint-active="yes"]').waypoint(function () {
            $(this).trigger('on-appear');
        }, { offset: '90%' });

        $(window).on('load', function () {
            setTimeout(function () {
                $.waypoints('refresh');
            }, 100);
        });
        $(window).load(function () {
            lastestTweets();
            googleMap();
            flickrFeed();
            testiCarousel();
            blogSlider();
        })
    });

})(jQuery);

$(document).ready(function () {

    $(".imgBanner").css("height", $(window).innerHeight() + "px");
    $(".imgBanner").css("top", "0px");
    //  container-empty

    $(window).resize(function () {
        $(".imgBanner").css("height", $(window).innerHeight() + "px");
        $(".imgBanner").css("top", "0px");
        $(".flex-control-nav").css("top", ($(window).innerHeight() - 80 - $(".contact-slider").height()) + "px");
        // alert($(window).innerHeight()  - $(".contact-slider").height());
        $(".flex-control-nav").css("z-index", "2");
        $(".control-slider").css("margin-top", "0px");


        // if ($(window).innerHeight() < 768) {
        //     $(".container-empty").css("height", "0px");
        // } else {
        //     $(".container-empty").css("height", "0px");
        // }
    });

    // alert($(window).innerHeight())
    //  var t = $(this).data("flexslider");

    // alert($.flexslider.data("flexslider")+asd);
    //switch (e) {
    //    case "play": t.play(); break;
    //    case "pause": t.pause(); break;
    //    case "stop": t.stop(); break;
    //    case "next": t.flexAnimate(t.getTarget("next"), !0); break;
    //    case "prev": case "previous": t.flexAnimate(t.getTarget("prev"), !0); break;
    //    default: "number" == typeof e && t.flexAnimate(e, !0)
    //}
});


