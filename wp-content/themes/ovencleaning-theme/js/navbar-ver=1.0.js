( function( $ ) {
	// Keyboard navigable dropdowns.
	$( '.js-cascade-navbar' ).find( 'a' ).on( 'focus blur', function() {
		$( this ).parents().toggleClass( 'focus' );
	} );

	// Touch-friendly dropdowns.
	if( ( 'ontouchstart' in window ) || ( navigator.msMaxTouchPoints > 0 ) || ( navigator.msMaxTouchPoints > 0 ) ) {
		$( '.js-cascade-navbar .has-children' ).each( function() {
			var currentItem = false;

			$( this ).on( 'click', function( e ) {
				var item = $( this ),
				    navbar = $( this ).closest( '.js-cascade-navbar' ),
				    toggle = $( '[aria-controls="' + navbar.attr('id') + '"]' );

				if ( item[0] !== currentItem[0] && toggle.css( 'display' ) === 'none' ) {
					e.preventDefault();
					currentItem = item;
				}
			} );

			$( document ).on( 'click touchstart MSPointerDown', function( e ) {
				var resetItem = true,
				    parents = $( e.target ).parents();

				for ( var i = 0; i < parents.length; i++ ) {
					if ( parents[i] === currentItem[0] ) {
						resetItem = false;
					}
				}

				if ( resetItem ) {
					currentItem = false;
				}
			} );
		} );
	}
} )( jQuery );