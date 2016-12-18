module VagrantVbguest
  module Installers
    class Windows < Base
      include VagrantVbguest::Helpers::VmCompatible

      def self.match?(vm)
        :windows == self.distro(vm)
      end

      def running?
        true
      end

      def install(opts=nil, &block)
        opts = opts || {}
        env.ui.warn "Using experimental Windows support"

        #upload(iso_file)
        #opts[:drive_letter] = mount_iso(opts, &block)
        opts[:drive_letter] = "D"
        env.ui.info "Got drive letter #{opts[:drive_letter]}"
        execute_installer(opts, &block)
        #unmount_iso(opts, &block) unless options[:no_cleanup]
      end

      def tmp_path
        options[:iso_upload_path] || 'C:\\VBoxGuestAdditions.iso'
      end

      def mount_iso(opts=null, &block)
        drive_letter = ""
        communicate.sudo("(Mount-DiskImage -ImagePath #{tmp_path} -StorageType ISO -Access ReadOnly -PassThru | Get-Volume).DriveLetter") do |type, data|
          drive_letter += data if type == :stdout
        end

        drive_letter.strip
      end

      def execute_installer(opts, &block)
        stream_output = Proc.new do |type, data|
          color = type == :stdout ? :green : :red
          env.ui.info data, color: color
        end

        install_cert = <<-SHELL
          $numCerts = (Get-ChildItem Cert:\\LocalMachine\\TrustedPublisher `
              | Where-Object { $_.Subject -like "*CN=Oracle Corporation, OU=VirtualBox*" }).Count

          if (!$numCerts) {
            & #{opts[:drive_letter]}:\\cert\\VBoxCertUtil.exe add-trusted-publisher D:\\cert\\oracle-vbox.cer
          }
        SHELL

        communicate.sudo(install_cert, &stream_output)
        communicate.sudo("& #{opts[:drive_letter]}:\\VBoxWindowsAdditions.exe /S", &stream_output)
        communicate.sudo("Restart-Computer -Force")
      end
    end
  end
end
VagrantVbguest::Installer.register(VagrantVbguest::Installers::Windows, 6)
