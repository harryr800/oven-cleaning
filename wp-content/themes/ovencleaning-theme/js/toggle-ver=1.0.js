( function( $ ) {
	/**
	 * Store button label for use when re-opening.
	 */
	$( '.js-cascade-toggle' ).attr( 'data-label-open', function() {
		return $(this).text();
	} );

	/**
	 * When clicked, toggle aria-expanded attribute on self, and toggle is-open
	 * class on target of aria-controls attribute.
	 */
	$( '.js-cascade-toggle' ).click( function() {
		var toggle = $( this );
		var target = $( '#' + toggle.attr( 'aria-controls' ) );
		var expanded = target.hasClass( 'is-open' );

		target.toggleClass( 'is-open' );

		expanded = target.hasClass( 'is-open' );

		toggle.attr( 'aria-expanded', expanded );

		toggle.text( expanded ? toggle.data( 'label-close' ) : toggle.data( 'label-open' ) );
	} );
} )( jQuery );