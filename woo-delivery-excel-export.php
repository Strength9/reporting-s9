<?php
/**
 * Plugin Name: WooCommerce Delivery Excel Export
 * Plugin URI: https://strength9.co.uk
 * Description: Export WooCommerce delivery orders to Excel spreadsheets  with customizable fields, date ranges, and order statuses. Simplify order management and delivery tracking.
 * Version: 1.0.0
 * Author: Dave Pratt
 * Author URI: https://strength9.co.uk
 * Text Domain: s9rp-cod
 * Domain Path: /languages
 * License: GPL v2 or later
 * License URI: https://www.gnu.org/licenses/gpl-2.0.html
 * Copyright: © 
 * 
 * Requires at least: 6.0
 * Requires PHP: 7.4
 */

// Prevent direct access to this file
if (!defined('ABSPATH')) {
    exit('Direct access not permitted.');
}

// Add security headers
add_action('init', function() {
    if (is_admin()) {
        header('X-Content-Type-Options: nosniff');
        header('X-Frame-Options: SAMEORIGIN');
        header('X-XSS-Protection: 1; mode=block');
    }
});

// Check if WooCommerce is active
function wde_check_woocommerce() {
    if (!in_array('woocommerce/woocommerce.php', apply_filters('active_plugins', get_option('active_plugins')))) {
        add_action('admin_notices', 'wde_woocommerce_missing_notice');
        return false;
    }
    return true;
}

// Admin notice if WooCommerce is not active
function wde_woocommerce_missing_notice() {
    ?>
    <div class="error">
        <p><?php esc_html_e('WooCommerce Delivery Excel Export requires WooCommerce to be installed and active.', 's9rp-cod'); ?></p>
    </div>
    <?php
}

// Initialize plugin
add_action('plugins_loaded', function() {
    if (wde_check_woocommerce()) {
        add_action('admin_menu', 'wde_add_menu_page');
        add_action('admin_enqueue_scripts', 'wde_admin_scripts');
        add_action('admin_init', 'wde_handle_export');
    }
});

// Add menu page
function wde_add_menu_page() {
    add_submenu_page(
        'woocommerce',
        __('Customer Contact Info Export', 's9rp-cod'),
        __('Customer Contact Info Export', 's9rp-cod'),
        'manage_woocommerce',
        'woo-delivery-export',
        'wde_admin_page'
    );
}

// Get orders data with enhanced security
function wde_get_orders_data($start_date, $end_date, $order_statuses, $parent_categories = array()) {
    // Validate inputs
    if (!current_user_can('manage_woocommerce')) {
        return array();
    }

    // Validate dates
    if (!wde_validate_date($start_date) || !wde_validate_date($end_date)) {
        return array();
    }

    // Validate order statuses
    $valid_statuses = array_keys(wc_get_order_statuses());
    $order_statuses = array_intersect($order_statuses, $valid_statuses);

    // Validate categories
    if (!empty($parent_categories)) {
        $parent_categories = array_filter($parent_categories, function($cat_id) {
            return term_exists($cat_id, 'product_cat');
        });
    }

    $args = array(
        'limit' => -1,
        'type' => 'shop_order',
        'status' => $order_statuses,
        'date_created' => $start_date . '...' . $end_date
    );

    $orders = wc_get_orders($args);
    $orders_data = array();

    // If no parent categories selected, return all orders
    if (empty($parent_categories)) {
        foreach ($orders as $order) {
            $orders_data[] = prepare_order_data($order);
        }
        return $orders_data;
    }

    // Build category tree (parent categories and their children)
    $category_tree = array();
    foreach ($parent_categories as $parent_id) {
        $category_tree[] = (int)$parent_id; // Add parent
        $children = get_term_children($parent_id, 'product_cat');
        if (!is_wp_error($children)) {
            foreach ($children as $child) {
                $category_tree[] = (int)$child;
            }
        }
    }
    $category_tree = array_unique($category_tree);

    // Process orders
    foreach ($orders as $order) {
        $include_order = false;

        // Check each item in the order
        foreach ($order->get_items() as $item) {
            $product = $item->get_product();
            if (!$product) continue;

            // If this is a variation, get the parent product for categories
            if ($product->is_type('variation')) {
                $parent_id = $product->get_parent_id();
                $product_categories = wp_get_post_terms($parent_id, 'product_cat', array('fields' => 'ids'));
            } else {
                $product_categories = wp_get_post_terms($product->get_id(), 'product_cat', array('fields' => 'ids'));
            }
            
            // If product has any category that's in our tree, include the order
            if (array_intersect($category_tree, $product_categories)) {
                $include_order = true;
                break; // Found a matching product, no need to check others
            }
        }

        if ($include_order) {
            $orders_data[] = prepare_order_data($order);
        }
    }

    return $orders_data;
}

// Validate date format
function wde_validate_date($date) {
    $d = DateTime::createFromFormat('Y-m-d', $date);
    return $d && $d->format('Y-m-d') === $date;
}

// Handle the export process with enhanced security
function wde_handle_export() {
    if (!isset($_POST['wde_export']) || !isset($_POST['wde_export_nonce'])) {
        return;
    }

    // Verify nonce and capabilities
    if (!wp_verify_nonce($_POST['wde_export_nonce'], 'wde_export_action') || 
        !current_user_can('manage_woocommerce')) {
        wp_die(
            esc_html__('Security check failed.', 's9rp-cod'),
            esc_html__('Error', 's9rp-cod'),
            array('response' => 403)
        );
    }

    // Validate and sanitize input
    $start_date = isset($_POST['start_date']) ? sanitize_text_field($_POST['start_date']) : '';
    $end_date = isset($_POST['end_date']) ? sanitize_text_field($_POST['end_date']) : '';
    
    if (!wde_validate_date($start_date) || !wde_validate_date($end_date)) {
        wp_die(
            esc_html__('Invalid date format.', 's9rp-cod'),
            esc_html__('Error', 's9rp-cod'),
            array('response' => 400)
        );
    }

    $order_statuses = isset($_POST['order_status']) ? 
        array_map('sanitize_text_field', $_POST['order_status']) : 
        array('wc-processing', 'wc-completed');

    $parent_categories = isset($_POST['parent_categories']) ? 
        array_map('absint', (array)$_POST['parent_categories']) : 
        array();

    // Get filtered orders
    $orders_data = wde_get_orders_data($start_date, $end_date, $order_statuses, $parent_categories);

    // Set secure headers for download
    nocache_headers();
    header('Content-Type: text/csv; charset=utf-8');
    header('Content-Disposition: attachment; filename="woocommerce-orders-export-' . date('Y-m-d') . '.csv"');
    header('Pragma: no-cache');
    header('Expires: 0');

    // Create output stream
    $output = fopen('php://output', 'w');

    // Add UTF-8 BOM for proper Excel encoding
    fputs($output, "\xEF\xBB\xBF");

    // Add headers
    fputcsv($output, array(
        'Order ID',
        'Order Date',
        'First Name',
        'Last Name',
        'Email Address',
        'Pickup Date',
        'Pickup Time'
    ));

    // Add order data
    foreach ($orders_data as $order) {
        fputcsv($output, array(
            $order['id'],
            $order['date'],
            $order['first_name'],
            $order['last_name'],
            $order['email'],
            $order['pickup_date'],
            $order['pickup_time']
        ));
    }

    fclose($output);
    exit;
}

// Enqueue admin scripts
function wde_admin_scripts($hook) {
    if ('woocommerce_page_woo-delivery-export' !== $hook) {
        return;
    }
    wp_enqueue_script('jquery-ui-datepicker');
    wp_enqueue_style('jquery-ui-datepicker', '//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css');
    
    // Add Select2 for better category selection
    wp_enqueue_script('select2', 'https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js', array('jquery'), '4.1.0', true);
    wp_enqueue_style('select2', 'https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css');
    
    // Add custom styles
    wp_add_inline_style('select2', '
        .wde-results-table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        .wde-results-table th, .wde-results-table td { padding: 8px; border: 1px solid #ddd; }
        .wde-results-table th { background-color: #f5f5f5; }
        .wde-results-table tr:nth-child(even) { background-color: #f9f9f9; }
        .wde-results-table tr:hover { background-color: #f5f5f5; }
        .select2-container { min-width: 300px; }
    ');
}

// Helper function to prepare order data
function prepare_order_data($order) {
    $order_id = $order->get_id();
    $order_link = admin_url('post.php?post=' . $order_id . '&action=edit');
    
    $pickup_date = get_post_meta($order_id, 'pickup_date', true);
    $pickup_time = get_post_meta($order_id, 'pickup_time', true);
    
    $order_date = $order->get_date_created()->date('d/m/Y H:i:s');
    $formatted_pickup_date = !empty($pickup_date) ? date('d/m/Y', strtotime($pickup_date)) : '';
    
    return array(
        'id' => $order_id,
        'date' => $order_date,
        'link' => $order_link,
        'first_name' => $order->get_billing_first_name(),
        'last_name' => $order->get_billing_last_name(),
        'email' => $order->get_billing_email(),
        'pickup_date' => $formatted_pickup_date,
        'pickup_time' => $pickup_time
    );
}

// Admin page content
function wde_admin_page() {
    $order_statuses = wc_get_order_statuses();
    $orders_data = array();
    $show_results = false;
    $parent_categories = array();
    $start_date = '';
    $end_date = '';
    $selected_statuses = array('wc-processing', 'wc-completed'); // Default statuses

    // Get only parent product categories
    $parent_cats = get_terms(array(
        'taxonomy' => 'product_cat',
        'hide_empty' => true,
        'parent' => 0 // Only get parent categories
    ));

    if (isset($_POST['wde_search']) && isset($_POST['wde_export_nonce']) && 
        wp_verify_nonce($_POST['wde_export_nonce'], 'wde_export_action')) {
        
        $start_date = sanitize_text_field($_POST['start_date']);
        $end_date = sanitize_text_field($_POST['end_date']);
        $selected_statuses = isset($_POST['order_status']) ? array_map('sanitize_text_field', $_POST['order_status']) : array('wc-processing', 'wc-completed');
        $parent_categories = isset($_POST['parent_categories']) ? array_map('absint', $_POST['parent_categories']) : array();
        
        $orders_data = wde_get_orders_data($start_date, $end_date, $selected_statuses, $parent_categories);
        $show_results = true;
    }
    ?>
    <div class="wrap">
        <h1><?php echo esc_html(get_admin_page_title()); ?></h1>
        <form method="post" action="">
            <?php wp_nonce_field('wde_export_action', 'wde_export_nonce'); ?>
            
            <table class="form-table">
                <tr>
                    <th scope="row"><?php _e('Date Range', 's9rp-cod'); ?></th>
                    <td>
                        <input type="text" id="start_date" name="start_date" class="datepicker" 
                            value="<?php echo esc_attr($start_date); ?>" placeholder="Start Date" required>
                        <input type="text" id="end_date" name="end_date" class="datepicker" 
                            value="<?php echo esc_attr($end_date); ?>" placeholder="End Date" required>
                    </td>
                </tr>
                <tr>
                    <th scope="row"><?php _e('Order Status', 's9rp-cod'); ?></th>
                    <td>
                        <?php foreach ($order_statuses as $status => $label) : ?>
                            <label>
                                <input type="checkbox" name="order_status[]" value="<?php echo esc_attr($status); ?>" 
                                    <?php checked(in_array($status, $selected_statuses)); ?>>
                                <?php echo esc_html($label); ?>
                            </label><br>
                        <?php endforeach; ?>
                    </td>
                </tr>
                <tr>
                    <th scope="row"><?php _e('Main Product Categories', 's9rp-cod'); ?></th>
                    <td>
                        <select name="parent_categories[]" id="parent_categories" multiple="multiple" class="select2">
                            <?php foreach ($parent_cats as $category) : ?>
                                <option value="<?php echo esc_attr($category->term_id); ?>"
                                    <?php selected(in_array($category->term_id, $parent_categories)); ?>>
                                    <?php echo esc_html($category->name); ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                        <p class="description"><?php _e('Optional: Filter orders by main product categories. Leave empty to show all orders.<br> <strong>Note:</strong> If Used this will only show orders that have products in the selected categories or their children.', 's9rp-cod'); ?></p>
                    </td>
                </tr>
            </table>

            <p class="submit">
                <button type="submit" name="wde_search" class="button button-primary">
                    <?php _e('Search Orders', 's9rp-cod'); ?>
                </button>
            </p>
        </form>

        <?php if ($show_results) : ?>
            <h2><?php _e('Search Results', 's9rp-cod'); ?></h2>
            
            <?php if (empty($orders_data)) : ?>
                <div class="notice notice-warning">
                    <p><?php _e('No orders found matching your search criteria.', 's9rp-cod'); ?></p>
                </div>
            <?php else : ?>
                <!-- Display the results table -->
                <table class="wde-results-table">
                    <thead>
                        <tr>
                            <th><?php _e('Order ID', 's9rp-cod'); ?></th>
                            <th><?php _e('Order Date', 's9rp-cod'); ?></th>
                            <th><?php _e('First Name', 's9rp-cod'); ?></th>
                            <th><?php _e('Last Name', 's9rp-cod'); ?></th>
                            <th><?php _e('Email Address', 's9rp-cod'); ?></th>
                            <th><?php _e('Pickup Date', 's9rp-cod'); ?></th>
                            <th><?php _e('Pickup Time', 's9rp-cod'); ?></th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php foreach ($orders_data as $order) : ?>
                            <tr>
                                <td><a href="<?php echo esc_url($order['link']); ?>"><?php echo esc_html($order['id']); ?></a></td>
                                <td><?php echo esc_html($order['date']); ?></td>
                                <td><?php echo esc_html($order['first_name']); ?></td>
                                <td><?php echo esc_html($order['last_name']); ?></td>
                                <td><?php echo esc_html($order['email']); ?></td>
                                <td><?php echo esc_html($order['pickup_date']); ?></td>
                                <td><?php echo esc_html($order['pickup_time']); ?></td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>

                <!-- Export form with hidden fields -->
                <form method="post" action="">
                    <?php wp_nonce_field('wde_export_action', 'wde_export_nonce'); ?>
                    <input type="hidden" name="start_date" value="<?php echo esc_attr($start_date); ?>">
                    <input type="hidden" name="end_date" value="<?php echo esc_attr($end_date); ?>">
                    <?php 
                    // Include selected categories in hidden fields
                    if (!empty($parent_categories)) {
                        foreach ($parent_categories as $cat_id) {
                            echo '<input type="hidden" name="parent_categories[]" value="' . esc_attr($cat_id) . '">';
                        }
                    }
                    // Include selected order statuses
                    if (!empty($selected_statuses)) {
                        foreach ($selected_statuses as $status) {
                            echo '<input type="hidden" name="order_status[]" value="' . esc_attr($status) . '">';
                        }
                    }
                    ?>
                    <p class="submit">
                        <button type="submit" name="wde_export" class="button button-primary">
                            <?php _e('Download as CSV', 's9rp-cod'); ?>
                        </button>
                    </p>
                </form>
            <?php endif; ?>
        <?php endif; ?>
    </div>

    <script>
    jQuery(document).ready(function($) {
        $('.datepicker').datepicker({
            dateFormat: 'yy-mm-dd',
            maxDate: '0'
        });

        $('.select2').select2({
            placeholder: '<?php _e('Select main categories (optional)', 's9rp-cod'); ?>',
            allowClear: true
        });
    });
    </script>
    <?php
} 