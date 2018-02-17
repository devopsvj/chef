#
# Author:: Doug Ireton (<doug@1strategy.com>)
# Copyright:: 2012-2018, Nordstrom, Inc.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# See here for more info:
# http://msdn.microsoft.com/en-us/library/windows/desktop/aa394492(v=vs.85).aspx

require "chef/resource"

class Chef
  class Resource
    # Create Windows printer. Note that this doesn't currently install a printer driver.
    # You must already have the driver installed on the system.
    class WindowsPrinter < Chef::Resource
      resource_name :windows_printer
      provides :windows_printer

      require "resolv"

      property :device_id, String, name_property: true
      property :comment, String
      property :default, [true, false], default: false
      property :driver_name, String, required: true
      property :location, String
      property :shared, [true, false], default: false
      property :share_name, String
      property :ipv4_address, String, regex: Resolv::IPv4::Regex
      property :exists, [true, false], desired_state: true

      PRINTERS_REG_KEY = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\\'.freeze unless defined?(PRINTERS_REG_KEY)

      # does the printer exist
      #
      # @param [String] name the name of the printer
      # @return [Boolean]
      def printer_exists?(name)
        printer_reg_key = PRINTERS_REG_KEY + name
        Chef::Log.debug "Checking to see if this reg key exists: '#{printer_reg_key}'"
        registry_key_exists?(printer_reg_key)
      end

      load_current_value do |desired|
        name desired.name
        exists printer_exists?(desired.name)
        # TODO: Set @current_resource printer properties from registry
      end

      action :create do
        if @current_resource.exists
          Chef::Log.info "#{@new_resource} already exists - nothing to do."
        else
          converge_by("Create #{@new_resource}") do
            create_printer
          end
        end
      end

      action :delete do
        if @current_resource.exists
          converge_by("Delete #{@new_resource}") do
            delete_printer
          end
        else
          Chef::Log.info "#{@current_resource} doesn't exist - can't delete."
        end
      end

      action_class do
        # creates the printer port and then the printer
        def create_printer
          # Create the printer port first
          windows_printer_port new_resource.ipv4_address do
          end

          port_name = "IP_#{new_resource.ipv4_address}"

          declare_resource(:powershell_script, "Creating printer: #{new_resource.name}") do
            code <<-EOH

              Set-WmiInstance -class Win32_Printer `
                -EnableAllPrivileges `
                -Argument @{ DeviceID   = "#{new_resource.device_id}";
                            Comment    = "#{new_resource.comment}";
                            Default    = "$#{new_resource.default}";
                            DriverName = "#{new_resource.driver_name}";
                            Location   = "#{new_resource.location}";
                            PortName   = "#{port_name}";
                            Shared     = "$#{new_resource.shared}";
                            ShareName  = "#{new_resource.share_name}";
                          }
            EOH
          end
        end

        def delete_printer
          declare_resource(:powershell_script, "Deleting printer: #{new_resource.name}") do
            code <<-EOH
              $printer = Get-WMIObject -class Win32_Printer -EnableAllPrivileges -Filter "name = '#{new_resource.name}'"
              $printer.Delete()
            EOH
          end
        end
      end
    end
  end
end
